import ssl
import requests
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util.ssl_ import create_urllib3_context
from zeep import Client
from zeep.transports import Transport

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

ALICUOTA_ID = {0: 3, 10.5: 4, 21: 5, 27: 6}
TIPO_FC     = [11, 12, 13]   # Factura/Nota C (monotributistas, sin IVA)

_clients = {}


class _LegacyTLSAdapter(HTTPAdapter):
    """Permite DH keys pequeñas que usan los servidores de AFIP."""
    def init_poolmanager(self, *args, **kwargs):
        ctx = create_urllib3_context()
        ctx.set_ciphers('DEFAULT:@SECLEVEL=1')
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        kwargs['ssl_context'] = ctx
        return super().init_poolmanager(*args, **kwargs)


def get_client(wsdl):
    if wsdl not in _clients:
        session = requests.Session()
        session.verify = False
        session.mount('https://', _LegacyTLSAdapter())
        transport = Transport(session=session, timeout=30)
        _clients[wsdl] = Client(wsdl, transport=transport)
    return _clients[wsdl]


def get_ultimo_comprobante(client, auth, punto_venta, tipo_cbte):
    result = client.service.FECompUltimoAutorizado(
        Auth=auth, PtoVta=punto_venta, CbteTipo=tipo_cbte
    )
    if result.Errors:
        for err in result.Errors.Err:
            if err.Code == 10:   # sin comprobantes emitidos aún
                return 0
        raise Exception(f"WSFE error: {result.Errors.Err[0].Msg}")
    return result.CbteNro


def procesar_comprobante(client, auth, cuit, punto_venta, tipo_cbte, comp, nro, cbtes_asoc=None):
    alicuota  = float(comp.get('alicuota', 21))
    imp_neto  = round(float(comp['imp_neto']), 2)
    imp_iva   = round(float(comp['imp_iva']),  2)
    imp_total = round(float(comp['imp_total']), 2)

    # Para FC el monto va en ImpNeto; para alicuota=0 en ImpOpEx
    if tipo_cbte in TIPO_FC:
        imp_op_ex = 0
    elif alicuota == 0:
        imp_op_ex = imp_neto
        imp_neto  = 0
    else:
        imp_op_ex = 0

    # Condición IVA receptor (RG 5616): 1=RI, 4=Exento, 5=CF, 6=Monotributo, 7=No categorizado
    condicion_iva = int(comp.get('condicion_iva', 0))
    if condicion_iva == 0:
        doc_tipo = int(comp['doc_tipo'])
        if doc_tipo == 99:
            condicion_iva = 5   # Consumidor Final
        elif doc_tipo == 80:
            condicion_iva = 1   # CUIT → Responsable Inscripto
        else:
            condicion_iva = 5   # default CF

    concepto = int(comp['concepto'])
    fecha    = str(comp['fecha'])  # YYYYMMDD

    det = {
        'Concepto': concepto,
        'DocTipo':  int(comp['doc_tipo']),
        'DocNro':   int(str(comp['doc_nro']).split('.')[0] or 0),
        'CbteDesde': nro,
        'CbteHasta': nro,
        'CbteFch':   fecha,
        'ImpTotal':   imp_total,
        'ImpTotConc': 0,
        'ImpNeto':    imp_neto,
        'ImpOpEx':    imp_op_ex,
        'ImpIVA':     imp_iva,
        'ImpTrib':    0,
        'MonId':      'PES',
        'MonCotiz':   1,
        'CondicionIVAReceptorId': condicion_iva,
    }

    # Fechas obligatorias para Servicios (concepto 2 o 3)
    if concepto in [2, 3]:
        det['FchServDesde'] = fecha
        det['FchServHasta'] = fecha
        det['FchVtoPago']   = fecha

    # Sección IVA solo para FA/FB con alícuota > 0
    if tipo_cbte not in TIPO_FC and alicuota > 0 and alicuota in ALICUOTA_ID:
        det['Iva'] = {
            'AlicIva': [{'Id': ALICUOTA_ID[alicuota], 'BaseImp': imp_neto, 'Importe': imp_iva}]
        }

    if cbtes_asoc:
        det['CbtesAsoc'] = {'CbteAsoc': [
            {'Tipo': a['tipo'], 'PtoVta': a['pv'], 'Nro': a['nro'], 'Cuit': int(cuit)}
            for a in cbtes_asoc
        ]}

    req = {
        'FeCabReq': {'CantReg': 1, 'PtoVta': punto_venta, 'CbteTipo': tipo_cbte},
        'FeDetReq': {'FECAEDetRequest': [det]},
    }

    return client.service.FECAESolicitar(Auth=auth, FeCAEReq=req)
