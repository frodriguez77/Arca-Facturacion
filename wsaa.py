import base64
import datetime
import email.utils
import json
import os
import subprocess
import tempfile
import requests
import xml.etree.ElementTree as ET
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from openssl_util import encontrar_openssl

_cache = {}
_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'wsaa_cache.json')


def _load_cache():
    if os.path.exists(_CACHE_FILE):
        try:
            with open(_CACHE_FILE, 'r') as f:
                data = json.load(f)
            for k, v in data.items():
                v['expiry'] = datetime.datetime.fromisoformat(v['expiry'])
                _cache[k] = v
        except Exception:
            pass

def _save_cache():
    data = {k: {**v, 'expiry': v['expiry'].isoformat()} for k, v in _cache.items()}
    with open(_CACHE_FILE, 'w') as f:
        json.dump(data, f)

_load_cache()


def _find_tag(elem, tag):
    for e in elem.iter():
        local = e.tag.split('}')[-1] if '}' in e.tag else e.tag
        if local == tag:
            return e
    return None


def _get_accurate_time():
    """Obtener hora Argentina precisa, corrigiendo desfasaje del reloj local."""
    tz_arg = datetime.timezone(datetime.timedelta(hours=-3))
    local_now = datetime.datetime.now(tz_arg)

    for url in ['https://www.google.com', 'https://www.microsoft.com',
                'https://www.cloudflare.com']:
        try:
            resp = requests.head(url, timeout=5, verify=False)
            date_str = resp.headers.get('Date')
            if date_str:
                server_utc = email.utils.parsedate_to_datetime(date_str)
                server_arg = server_utc.astimezone(tz_arg)
                diff = (server_arg - local_now).total_seconds()
                if abs(diff) > 30:
                    print(f"WSAA: Reloj local desfasado {diff:+.0f}s — "
                          f"usando hora del servidor ({url})")
                return server_arg
        except Exception:
            continue

    print("WSAA: No se pudo verificar hora online, usando reloj local")
    return local_now


def _generate_tra(service):
    now    = _get_accurate_time()
    expiry = now + datetime.timedelta(hours=12)
    fmt    = "%Y-%m-%dT%H:%M:%S-03:00"
    tra = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<loginTicketRequest version="1.0">\n'
        '  <header>\n'
        f'    <uniqueId>{int(now.timestamp())}</uniqueId>\n'
        f'    <generationTime>{now.strftime(fmt)}</generationTime>\n'
        f'    <expirationTime>{expiry.strftime(fmt)}</expirationTime>\n'
        '  </header>\n'
        f'  <service>{service}</service>\n'
        '</loginTicketRequest>\n'
    )
    print(f"TRA: generationTime={now.strftime(fmt)}")
    return tra.encode('utf-8')


def _sign_tra(tra_bytes, cert_path, key_path):
    tra_file = tempfile.mktemp(suffix='.xml')
    try:
        with open(tra_file, 'wb') as f:
            f.write(tra_bytes)

        result = subprocess.run(
            [encontrar_openssl(), 'smime', '-sign',
             '-in',     tra_file,
             '-signer', cert_path,
             '-inkey',  key_path,
             '-outform', 'DER',
             '-nodetach',
             '-md', 'sha256'],
            capture_output=True
        )

        if result.returncode != 0:
            raise Exception(f"Error OpenSSL: {result.stderr.decode('utf-8', errors='replace')}")

        der = result.stdout
        print(f"DER: {len(der)} bytes, inicio={der[:4].hex()}")
        return base64.b64encode(der).decode('ascii')
    finally:
        if os.path.exists(tra_file):
            os.unlink(tra_file)


def get_ticket(service, cert_path, key_path, wsaa_url, cuit):
    # Cache key includes CUIT so each empresa has its own token
    cache_key = f"{cuit}_{service}"
    cached = _cache.get(cache_key)
    if cached and datetime.datetime.now() < cached['expiry']:
        return cached['token'], cached['sign']

    tra     = _generate_tra(service)
    cms_b64 = _sign_tra(tra, cert_path, key_path)

    soap = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soapenv:Envelope '
        'xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:wsaa="http://wsaa.view.sua.dvadmin.afip.gov.ar/ws/services/LoginCms">'
        '<soapenv:Header/>'
        '<soapenv:Body>'
        '<wsaa:loginCms>'
        f'<wsaa:in0>{cms_b64}</wsaa:in0>'
        '</wsaa:loginCms>'
        '</soapenv:Body>'
        '</soapenv:Envelope>'
    )

    resp = requests.post(
        wsaa_url,
        data=soap.encode('utf-8'),
        headers={'Content-Type': 'text/xml; charset=UTF-8', 'SOAPAction': ''},
        verify=False,
        timeout=60,
    )

    if not resp.ok:
        if 'alreadyAuthenticated' in resp.text:
            raise Exception("AFIP: ya existe un ticket válido. Esperá unos minutos y reintentá.")
        if 'generationTime' in resp.text:
            _cache.pop(cache_key, None)
            _save_cache()
            raise Exception(
                "AFIP rechazó la hora del sistema. "
                "Verificá que la fecha y hora de tu PC sean correctas: "
                "Configuración → Hora e idioma → Activar 'Ajustar hora automáticamente'. "
                f"(Detalle: {resp.text[:400]})"
            )
        raise Exception(f"WSAA HTTP {resp.status_code}: {resp.text[:800]}")

    root    = ET.fromstring(resp.content)
    ta_elem = _find_tag(root, 'loginCmsReturn')
    if ta_elem is None:
        raise Exception(f"WSAA sin loginCmsReturn: {resp.text[:400]}")

    ta         = ET.fromstring(ta_elem.text)
    token      = _find_tag(ta, 'token').text
    sign       = _find_tag(ta, 'sign').text
    expiry_str = _find_tag(ta, 'expirationTime').text

    expiry_dt = datetime.datetime.fromisoformat(expiry_str[:19]) - datetime.timedelta(minutes=10)
    _cache[cache_key] = {'token': token, 'sign': sign, 'expiry': expiry_dt}
    _save_cache()

    return token, sign
