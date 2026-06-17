from wsfe import get_client

PADRON_WSDL_PROD = 'https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5?wsdl'
PADRON_WSDL_HOMO = 'https://awshomo.afip.gov.ar/sr-padron/webservices/personaServiceA5?wsdl'
PADRON_SERVICE   = 'ws_sr_constancia_inscripcion'


def consultar_persona(token: str, sign: str, cuit_rep: str, cuit_consulta: str, wsdl: str) -> dict:
    client = get_client(wsdl)
    result = client.service.getPersona(
        token=token,
        sign=sign,
        cuitRepresentada=int(cuit_rep),
        idPersona=int(cuit_consulta),
    )

    # ws_sr_constancia_inscripcion devuelve los datos directamente en result.
    # ws_sr_padron_a4 los envuelve en result.persona.
    if hasattr(result, 'persona'):
        if getattr(result, 'errorConstancia', None):
            errs = result.errorConstancia.error
            raise Exception(errs[0].descripcion if errs else 'Error en consulta padrón')
        p = result.persona
    else:
        p = result

    tipo = str(getattr(p, 'tipoPersona', '') or '')

    if tipo == 'FISICA':
        apellido     = str(getattr(p, 'apellido', '') or '').strip()
        nombre_pila  = str(getattr(p, 'nombre',   '') or '').strip()
        razon_social = f'{apellido} {nombre_pila}'.strip()
    else:
        razon_social = str(getattr(p, 'razonSocial', '') or '').strip()

    # ws_sr_constancia_inscripcion usa domicilioFiscal (objeto único).
    # ws_sr_padron_a4 usa domicilio (lista).
    domicilio = ''
    dom_fiscal = getattr(p, 'domicilioFiscal', None)
    if dom_fiscal:
        parts = [
            str(getattr(dom_fiscal, 'direccion',           '') or ''),
            str(getattr(dom_fiscal, 'localidad',            '') or ''),
            str(getattr(dom_fiscal, 'descripcionProvincia', '') or ''),
        ]
        domicilio = ', '.join(x for x in parts if x)
    else:
        domicilios = list(getattr(p, 'domicilio', None) or [])
        candidatos = [d for d in domicilios if str(getattr(d, 'tipoDomicilio', '')) == 'FISCAL']
        if not candidatos:
            candidatos = domicilios
        if candidatos:
            d = candidatos[0]
            parts = [
                str(getattr(d, 'direccion',            '') or ''),
                str(getattr(d, 'localidad',             '') or ''),
                str(getattr(d, 'descripcionProvincia',  '') or ''),
            ]
            domicilio = ', '.join(x for x in parts if x)

    return {
        'cuit':         cuit_consulta,
        'tipo':         tipo,
        'razon_social': razon_social,
        'domicilio':    domicilio,
        'estado':       str(getattr(p, 'estadoClave', '') or ''),
        'condicion_iva': _extraer_condicion_iva(p),
    }


def _extraer_condicion_iva(p) -> str:
    if getattr(p, 'datosMonotributo', None):
        return 'Responsable Monotributo'

    datos_rg = getattr(p, 'datosRegimenGeneral', None)
    if datos_rg:
        imp_list = list(getattr(datos_rg, 'impuesto', None) or [])
        ids_imp = set()
        for imp in imp_list:
            try:
                ids_imp.add(int(getattr(imp, 'idImpuesto', 0)))
            except (ValueError, TypeError):
                pass
        if 30 in ids_imp:
            return 'IVA Responsable Inscripto'
        if 32 in ids_imp:
            return 'IVA Sujeto Exento'
        return 'IVA Responsable Inscripto'

    return 'Consumidor Final'
