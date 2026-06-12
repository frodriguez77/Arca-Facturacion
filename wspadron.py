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

    if getattr(result, 'errorConstancia', None):
        errs = result.errorConstancia.error
        raise Exception(errs[0].descripcion if errs else 'Error en consulta padrón')

    p = result.persona
    tipo = str(getattr(p, 'tipoPersona', '') or '')

    if tipo == 'FISICA':
        apellido     = str(getattr(p, 'apellido', '') or '').strip()
        nombre_pila  = str(getattr(p, 'nombre',   '') or '').strip()
        razon_social = f'{apellido} {nombre_pila}'.strip()
    else:
        razon_social = str(getattr(p, 'razonSocial', '') or '').strip()

    domicilio = ''
    domicilios = list(getattr(p, 'domicilio', None) or [])
    candidatos = [d for d in domicilios if str(getattr(d, 'tipoDomicilio', '')) == 'FISCAL']
    if not candidatos:
        candidatos = domicilios
    if candidatos:
        d     = candidatos[0]
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
    }
