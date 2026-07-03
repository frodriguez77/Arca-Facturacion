import functools
import io
import json
import os
import re
import subprocess
import traceback
import uuid
import zipfile
from datetime import datetime, timedelta

from flask import (Flask, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

import config
import wsaa
import wsfe
import factura_pdf
from openssl_util import encontrar_openssl
from repository import EmpresaRepository, UsuarioRepository

app = Flask(__name__)
app.secret_key = 'arca_2026_secret_key_change_in_prod'

BASE        = os.path.dirname(os.path.abspath(__file__))
UPLOAD       = os.path.join(BASE, 'uploads')
CERTS        = os.path.join(BASE, 'certificados')
CLIENTES_DIR = os.path.join(BASE, 'clientes')
COMPRAS_DIR  = os.path.join(BASE, 'compras')
FOTOS_CLI    = os.path.join(BASE, 'static', 'clientes')

_version_path = os.path.join(BASE, 'VERSION')
APP_VERSION   = open(_version_path).read().strip() if os.path.exists(_version_path) else '—'

_UPDATE_URL   = 'https://raw.githubusercontent.com/frodriguez77/Arca-Facturacion/claude/new-pc-download-setup-5AF1K/VERSION'
_update_cache = {'latest': None, 'checked_at': None}

@app.context_processor
def _inject_version():
    return {'app_version': APP_VERSION}

COLUMNAS = [
    'punto_venta', 'tipo_cbte', 'concepto',
    'doc_tipo', 'doc_nro', 'razon_social',
    'fecha', 'imp_neto', 'alicuota', 'imp_iva', 'imp_total',
]

NC_TIPO_MAP = {
    1: 3,   2: 3,    # Factura/ND A  → NC A
    6: 8,   7: 8,    # Factura/ND B  → NC B
    11: 13, 12: 13,  # Factura/ND C  → NC C
    51: 53, 52: 53,  # Factura/ND M  → NC M
    201: 203, 202: 203,
    206: 208, 207: 208,
    211: 213, 212: 213,
}

ND_TIPO_MAP = {
    1: 2,             # Factura A  → ND A
    6: 7,             # Factura B  → ND B
    11: 12,           # Factura C  → ND C
    51: 52,           # Factura M  → ND M
    201: 202, 206: 207, 211: 212,
}

TIPO_NOMBRE = {
    1: 'Factura A',       2: 'Nota de Débito A',  3: 'Nota de Crédito A',
    6: 'Factura B',       7: 'Nota de Débito B',  8: 'Nota de Crédito B',
    11: 'Factura C',      12: 'Nota de Débito C', 13: 'Nota de Crédito C',
    51: 'Factura M',      52: 'Nota de Débito M', 53: 'Nota de Crédito M',
    201: 'FCE A',         203: 'NCE A',
    206: 'FCE B',         208: 'NCE B',
    211: 'FCE C',         213: 'NCE C',
}

os.makedirs(UPLOAD,       exist_ok=True)
os.makedirs(CERTS,        exist_ok=True)
os.makedirs(CLIENTES_DIR, exist_ok=True)
os.makedirs(COMPRAS_DIR,  exist_ok=True)
os.makedirs(FOTOS_CLI,    exist_ok=True)


def _str_cae(v) -> str:
    """Convierte CAE o fecha numérica (int/float/str de Excel) a string limpio sin decimales."""
    if not v and v != 0:
        return ''
    s = str(v).strip()
    if s.endswith('.0'):
        return s[:-2]
    return s


# ---------- helpers -----------------------------------------------------------

def _mes_actual() -> str:
    """Devuelve el mes actual en formato YYYY-MM."""
    return datetime.today().strftime('%Y-%m')

def _carpeta_empresa(cuit: str, mes: str | None = None) -> str:
    """
    Devuelve (y crea si no existe) la carpeta de uploads para una empresa y mes.
    Estructura: uploads/{CUIT}/{YYYY-MM}/
    """
    carpeta = os.path.join(UPLOAD, cuit, mes or _mes_actual())
    os.makedirs(carpeta, exist_ok=True)
    return carpeta

def _upload_path(empresa_id: str, mes: str | None = None) -> str:
    empresa = EmpresaRepository.get_by_id(empresa_id)
    cuit    = empresa['cuit'] if empresa else empresa_id
    return os.path.join(_carpeta_empresa(cuit, mes), 'facturas.xlsx')

def _resultado_path(empresa_id: str, mes: str | None = None) -> str:
    empresa = EmpresaRepository.get_by_id(empresa_id)
    cuit    = empresa['cuit'] if empresa else empresa_id
    return os.path.join(_carpeta_empresa(cuit, mes), 'facturas_resultado.xlsx')

def _upload_path_actual(empresa_id: str) -> str:
    """Ruta del archivo subido en el mes actual."""
    return _upload_path(empresa_id, _mes_actual())

def _resultado_path_actual(empresa_id: str) -> str:
    """Ruta del resultado en el mes actual."""
    return _resultado_path(empresa_id, _mes_actual())

def _empresa_urls(empresa: dict) -> tuple[str, str]:
    if empresa.get('homologacion'):
        return config.WSAA_URL_HOMO, config.WSFE_WSDL_HOMO
    return config.WSAA_URL_PROD, config.WSFE_WSDL_PROD

def _get_current_user() -> dict | None:
    uid = session.get('user_id')
    if not uid:
        return None
    return UsuarioRepository.get_by_id(uid)

def _user_empresas(user: dict) -> list[dict]:
    """Empresas accesibles para el usuario."""
    if user['rol'] == 'admin' or not user.get('empresas'):
        return EmpresaRepository.get_all()
    return EmpresaRepository.filter_by_ids(user.get('empresas', []))

def _user_can_access(user: dict, empresa_id: str) -> bool:
    """True si el usuario puede operar sobre esa empresa."""
    if user['rol'] == 'admin':
        return True
    return empresa_id in user.get('empresas', [])

def _unique_id(base: str, exists_fn) -> str:
    """Genera un ID único a partir de un nombre base."""
    candidate = re.sub(r'[^a-z0-9]', '', base.lower())[:20] or str(uuid.uuid4())[:8]
    if exists_fn(candidate):
        candidate = candidate + '_' + str(uuid.uuid4())[:4]
    return candidate


# ---------- helpers de base de clientes ---------------------------------------

def _clientes_path(empresa_id: str) -> str:
    return os.path.join(CLIENTES_DIR, f'{empresa_id}.json')

def _load_clientes(empresa_id: str) -> dict:
    """Devuelve dict {cuit: {cuit, nombre, domicilio, estado}}."""
    path = _clientes_path(empresa_id)
    if not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as f:
        lista = json.load(f)
    return {str(c['cuit']): c for c in lista}

def _save_clientes(empresa_id: str, clientes: dict):
    path = _clientes_path(empresa_id)
    lista = sorted(clientes.values(), key=lambda x: x.get('nombre', ''))
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)

def _get_cliente(empresa_id: str, doc_nro) -> dict | None:
    """Devuelve el cliente por CUIT/doc_nro, o None si no existe."""
    cuit = str(doc_nro).split('.')[0].strip()
    if not cuit or cuit in ('0', ''):
        return None
    return _load_clientes(empresa_id).get(cuit)

def _compras_path(empresa_id: str) -> str:
    return os.path.join(COMPRAS_DIR, f'{empresa_id}.json')

def _load_compras(empresa_id: str) -> list:
    path = _compras_path(empresa_id)
    if not os.path.exists(path):
        return []
    with open(path, encoding='utf-8') as f:
        return json.load(f)

def _save_compras(empresa_id: str, compras: list):
    path = _compras_path(empresa_id)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(compras, f, ensure_ascii=False, indent=2)


# ---------- inicialización ---------------------------------------------------

def _init_admin():
    """Crea el usuario admin por defecto si no hay ningún usuario registrado."""
    if UsuarioRepository.count() == 0:
        UsuarioRepository.add({
            'id':            'admin',
            'username':      'admin',
            'password_hash': generate_password_hash('admin123'),
            'rol':           'admin',
            'nombre':        'Administrador',
            'empresas':      [],
        })
        print("Usuario admin creado — contraseña: admin123 (cambiala desde Admin)")

_init_admin()


# ---------- decoradores de autenticación -------------------------------------

def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not _get_current_user():
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        user = _get_current_user()
        if not user:
            return redirect(url_for('login'))
        if user['rol'] != 'admin':
            return render_template('error.html',
                                   mensaje='No tenés permisos para acceder a esta página.'), 403
        return f(*args, **kwargs)
    return decorated


# ---------- login / logout ----------------------------------------------------

@app.route('/login', methods=['GET', 'POST'])
def login():
    if _get_current_user():
        return redirect(url_for('index'))

    error = None
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        usuario  = UsuarioRepository.get_by_username(username)
        if usuario and check_password_hash(usuario['password_hash'], password):
            session['user_id'] = usuario['id']
            return redirect(url_for('index'))
        error = 'Usuario o contraseña incorrectos.'

    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ---------- vistas principales -----------------------------------------------

@app.route('/')
@login_required
def index():
    user     = _get_current_user()
    empresas = _user_empresas(user)
    return render_template('index.html', empresas=empresas, current_user=user)


# ---------- Empresa activa en sesión ------------------------------------------

@app.route('/api/empresa-activa', methods=['POST'])
@login_required
def api_set_empresa_activa():
    data = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    user = _get_current_user()
    if empresa_id and _user_can_access(user, empresa_id):
        session['empresa_activa'] = empresa_id
        return jsonify({'ok': True})
    return jsonify({'error': 'Acceso denegado'}), 403

@app.route('/api/empresa-activa')
@login_required
def api_get_empresa_activa():
    return jsonify({'empresa_id': session.get('empresa_activa', '')})

# ---------- API empresas ------------------------------------------------------

@app.route('/api/empresas')
@login_required
def api_empresas():
    return jsonify(_user_empresas(_get_current_user()))

@app.route('/api/empresas', methods=['POST'])
@admin_required
def api_empresa_add():
    data   = request.get_json(force=True)
    nombre = (data.get('nombre') or '').strip()
    cuit   = (data.get('cuit')   or '').strip()
    cert   = (data.get('cert')   or '').strip()
    key    = (data.get('key')    or '').strip()

    if not nombre or not cuit:
        return jsonify({'error': 'Nombre y CUIT son obligatorios'}), 400
    if not re.fullmatch(r'\d{11}', cuit):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400
    if EmpresaRepository.cuit_exists(cuit):
        return jsonify({'error': f'Ya existe una empresa con CUIT {cuit}'}), 400

    empresa_id = _unique_id(nombre, EmpresaRepository.id_exists)
    EmpresaRepository.add({
        'id':                 empresa_id,
        'nombre':             nombre,
        'cuit':               cuit,
        'cert':               cert,
        'key':                key,
        'homologacion':       bool(data.get('homologacion', False)),
        'domicilio':          (data.get('domicilio') or '').strip(),
        'telefono':           (data.get('telefono')  or '').strip(),
        'localidad':          (data.get('localidad') or '').strip(),
        'ing_brutos':         (data.get('ing_brutos') or '').strip(),
        'inicio_actividades': (data.get('inicio_actividades') or '').strip(),
        'matricula':          (data.get('matricula') or '').strip(),
        'logo_path':          (data.get('logo_path') or '').strip(),
        'smtp_server':        (data.get('smtp_server')      or '').strip(),
        'smtp_port':          int(data.get('smtp_port')      or 587),
        'smtp_user':          (data.get('smtp_user')         or '').strip(),
        'smtp_password':      (data.get('smtp_password')     or '').strip(),
        'smtp_ssl':           bool(data.get('smtp_ssl', False)),
        'nombre_remitente':   (data.get('nombre_remitente')  or '').strip(),
    })
    return jsonify({'ok': True, 'id': empresa_id})

@app.route('/api/empresas/<empresa_id>', methods=['PUT'])
@admin_required
def api_empresa_edit(empresa_id):
    if not EmpresaRepository.get_by_id(empresa_id):
        return jsonify({'error': 'Empresa no encontrada'}), 404

    data   = request.get_json(force=True)
    nombre = (data.get('nombre') or '').strip()
    cuit   = (data.get('cuit')   or '').strip()
    cert   = (data.get('cert')   or '').strip()
    key    = (data.get('key')    or '').strip()

    if not nombre or not cuit:
        return jsonify({'error': 'Nombre y CUIT son obligatorios'}), 400
    if not re.fullmatch(r'\d{11}', cuit):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400
    if EmpresaRepository.cuit_exists(cuit, exclude_id=empresa_id):
        return jsonify({'error': f'Ya existe otra empresa con CUIT {cuit}'}), 400

    upd = {
        'nombre':             nombre,
        'cuit':               cuit,
        'cert':               cert,
        'key':                key,
        'homologacion':       bool(data.get('homologacion', False)),
        'domicilio':          (data.get('domicilio') or '').strip(),
        'telefono':           (data.get('telefono')  or '').strip(),
        'localidad':          (data.get('localidad') or '').strip(),
        'ing_brutos':         (data.get('ing_brutos') or '').strip(),
        'inicio_actividades': (data.get('inicio_actividades') or '').strip(),
        'matricula':          (data.get('matricula') or '').strip(),
        'logo_path':          (data.get('logo_path') or '').strip(),
        'smtp_server':        (data.get('smtp_server')      or '').strip(),
        'smtp_port':          int(data.get('smtp_port')      or 587),
        'smtp_user':          (data.get('smtp_user')         or '').strip(),
        'smtp_ssl':           bool(data.get('smtp_ssl', False)),
        'nombre_remitente':   (data.get('nombre_remitente')  or '').strip(),
    }
    new_pw = (data.get('smtp_password') or '').strip()
    if new_pw and '•' not in new_pw:
        upd['smtp_password'] = new_pw
    EmpresaRepository.update(empresa_id, upd)
    return jsonify({'ok': True})

@app.route('/api/empresas/<empresa_id>', methods=['DELETE'])
@admin_required
def api_empresa_delete(empresa_id):
    if not EmpresaRepository.delete(empresa_id):
        return jsonify({'error': 'Empresa no encontrada'}), 404
    for path in [_upload_path(empresa_id), _resultado_path(empresa_id)]:
        if os.path.exists(path):
            os.remove(path)
    return jsonify({'ok': True})

@app.route('/api/upload-logo', methods=['POST'])
@admin_required
def api_upload_logo():
    f = request.files.get('logo')
    if not f or not f.filename:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'):
        return jsonify({'error': 'Formato no válido (use PNG o JPG)'}), 400
    logos_dir = os.path.join(BASE, 'static', 'logos')
    os.makedirs(logos_dir, exist_ok=True)
    filename = str(uuid.uuid4()) + ext
    path = os.path.join(logos_dir, filename)
    f.save(path)
    return jsonify({'path': path, 'url': f'/static/logos/{filename}'})


@app.route('/api/upload-cert', methods=['POST'])
@admin_required
def api_upload_cert():
    empresa_id = (request.form.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

    cuit     = empresa['cuit']
    cert_dir = os.path.join(CERTS, cuit)
    os.makedirs(cert_dir, exist_ok=True)

    resultado = {}

    f_cert = request.files.get('cert_file')
    if f_cert and f_cert.filename:
        if not f_cert.filename.lower().endswith('.crt'):
            return jsonify({'error': 'El certificado debe tener extensión .crt'}), 400
        cert_path = os.path.join(cert_dir, f'{cuit}.crt')
        f_cert.save(cert_path)
        EmpresaRepository.update(empresa_id, {'cert': cert_path})
        resultado['cert_path'] = cert_path

    f_key = request.files.get('key_file')
    if f_key and f_key.filename:
        if not f_key.filename.lower().endswith('.key'):
            return jsonify({'error': 'La clave privada debe tener extensión .key'}), 400
        key_path = os.path.join(cert_dir, f'{cuit}_clave.key')
        f_key.save(key_path)
        EmpresaRepository.update(empresa_id, {'key': key_path})
        resultado['key_path'] = key_path

    if not resultado:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400

    return jsonify({'ok': True, **resultado})


@app.route('/api/logo-preview')
@login_required
def api_logo_preview():
    path = request.args.get('path', '')
    try:
        rp   = os.path.realpath(path)
        base = os.path.realpath(BASE)
        if not rp.startswith(base) or not os.path.isfile(rp):
            return '', 404
    except Exception:
        return '', 404
    return send_file(rp)


# ---------- API usuarios ------------------------------------------------------

@app.route('/api/usuarios')
@admin_required
def api_usuarios():
    return jsonify(UsuarioRepository.get_all_public())

@app.route('/api/usuarios', methods=['POST'])
@admin_required
def api_usuario_add():
    data     = request.get_json(force=True)
    username = (data.get('username') or '').strip()
    nombre   = (data.get('nombre')   or '').strip()
    password = (data.get('password') or '').strip()
    rol      = (data.get('rol')      or 'usuario').strip()
    empresas = data.get('empresas', [])

    if not username or not password:
        return jsonify({'error': 'Usuario y contraseña son obligatorios'}), 400
    if rol not in ('admin', 'usuario'):
        return jsonify({'error': 'Rol inválido'}), 400
    if UsuarioRepository.username_exists(username):
        return jsonify({'error': 'Ya existe un usuario con ese nombre'}), 400

    uid = _unique_id(username, UsuarioRepository.id_exists)
    UsuarioRepository.add({
        'id':            uid,
        'username':      username,
        'password_hash': generate_password_hash(password),
        'rol':           rol,
        'nombre':        nombre or username,
        'empresas':      empresas if rol == 'usuario' else [],
    })
    return jsonify({'ok': True, 'id': uid})

@app.route('/api/usuarios/<uid>', methods=['PUT'])
@admin_required
def api_usuario_edit(uid):
    if not UsuarioRepository.get_by_id(uid):
        return jsonify({'error': 'Usuario no encontrado'}), 404

    data     = request.get_json(force=True)
    username = (data.get('username') or '').strip()
    nombre   = (data.get('nombre')   or '').strip()
    password = (data.get('password') or '').strip()
    rol      = (data.get('rol')      or 'usuario').strip()
    empresas = data.get('empresas', [])

    if not username:
        return jsonify({'error': 'El nombre de usuario es obligatorio'}), 400
    if rol not in ('admin', 'usuario'):
        return jsonify({'error': 'Rol inválido'}), 400
    if UsuarioRepository.username_exists(username, exclude_id=uid):
        return jsonify({'error': 'Ese nombre de usuario ya está en uso'}), 400

    fields = {
        'username': username,
        'nombre':   nombre or username,
        'rol':      rol,
        'empresas': empresas if rol == 'usuario' else [],
    }
    if password:
        fields['password_hash'] = generate_password_hash(password)

    UsuarioRepository.update(uid, fields)
    return jsonify({'ok': True})

@app.route('/api/usuarios/<uid>', methods=['DELETE'])
@admin_required
def api_usuario_delete(uid):
    user = _get_current_user()
    if user['id'] == uid:
        return jsonify({'error': 'No podés eliminar tu propio usuario'}), 400
    if not UsuarioRepository.delete(uid):
        return jsonify({'error': 'Usuario no encontrado'}), 404
    return jsonify({'ok': True})

@app.route('/api/usuarios/<uid>/cambiar-password', methods=['POST'])
@login_required
def api_cambiar_password(uid):
    user = _get_current_user()
    if user['id'] != uid and user['rol'] != 'admin':
        return jsonify({'error': 'Sin permisos'}), 403

    data         = request.get_json(force=True)
    nueva        = (data.get('nueva') or '').strip()
    confirmacion = (data.get('confirmacion') or '').strip()

    if not nueva or len(nueva) < 6:
        return jsonify({'error': 'La contraseña debe tener al menos 6 caracteres'}), 400
    if nueva != confirmacion:
        return jsonify({'error': 'Las contraseñas no coinciden'}), 400
    if not UsuarioRepository.get_by_id(uid):
        return jsonify({'error': 'Usuario no encontrado'}), 404

    UsuarioRepository.update(uid, {'password_hash': generate_password_hash(nueva)})
    return jsonify({'ok': True})


# ---------- flujo de facturación ----------------------------------------------

@app.route('/upload', methods=['POST'])
@login_required
def upload():
    user       = _get_current_user()
    empresa_id = request.form.get('empresa_id', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Seleccioná una empresa antes de cargar el archivo'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'No tenés acceso a esta empresa'}), 403

    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No se seleccionó archivo'}), 400
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    path = _upload_path_actual(empresa_id)
    f.save(path)

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
        faltantes = [c for c in COLUMNAS if c not in df.columns]
        if faltantes:
            return jsonify({'error': f'Faltan columnas: {", ".join(faltantes)}'}), 400

        CUIT_GENERICO = '20222222223'
        registros = []
        filas_cuit_insertado = []
        for idx, row in df.iterrows():
            r = {}
            for c in COLUMNAS:
                v = row[c]
                if hasattr(v, 'strftime'):
                    v = v.strftime('%Y-%m-%d')
                r[c] = str(v) if v is not None else ''
            r['forma_pago'] = str(row.get('forma_pago', '') or '').strip()

            doc_nro = r.get('doc_nro', '').replace('.0', '').strip()
            doc_tipo = r.get('doc_tipo', '').replace('.0', '').strip()
            if not doc_nro or doc_nro == '' or doc_nro == 'nan' or doc_nro == '0':
                r['doc_nro'] = CUIT_GENERICO
                if not doc_tipo or doc_tipo == '' or doc_tipo == 'nan' or doc_tipo == '0':
                    r['doc_tipo'] = '99'
                filas_cuit_insertado.append(idx + 2)

            registros.append(r)

        if filas_cuit_insertado:
            for fila_excel in filas_cuit_insertado:
                i = fila_excel - 2
                df.at[i, 'doc_nro'] = CUIT_GENERICO
                if str(df.at[i, 'doc_tipo']).strip() in ('', 'nan', '0'):
                    df.at[i, 'doc_tipo'] = 99
            df.to_excel(path, index=False)

        resp = {'ok': True, 'registros': registros, 'total': len(registros)}
        if filas_cuit_insertado:
            resp['cuit_generico_filas'] = filas_cuit_insertado
            resp['cuit_generico_msg'] = (
                f'Se insertó el CUIT genérico ({CUIT_GENERICO}) en '
                f'{len(filas_cuit_insertado)} fila(s) sin documento: '
                f'fila(s) {", ".join(str(f) for f in filas_cuit_insertado[:20])}'
                + ('...' if len(filas_cuit_insertado) > 20 else '')
            )
        return jsonify(resp)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/procesar', methods=['POST'])
@login_required
def procesar():
    user       = _get_current_user()
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'No tenés acceso a esta empresa'}), 403

    path = _upload_path_actual(empresa_id)
    if not os.path.exists(path):
        return jsonify({'error': 'No hay archivo cargado para esta empresa'}), 400

    cert_path = empresa.get('cert', '')
    key_path  = empresa.get('key', '')
    if not os.path.isfile(cert_path):
        return jsonify({'error': f'Certificado no encontrado: {cert_path}'}), 400
    if not os.path.isfile(key_path):
        return jsonify({'error': f'Clave privada no encontrada: {key_path}'}), 400

    wsaa_url, wsfe_wsdl = _empresa_urls(empresa)

    try:
        token, sign = wsaa.get_ticket('wsfe', cert_path, key_path, wsaa_url, empresa['cuit'])
        auth   = {'Token': token, 'Sign': sign, 'Cuit': int(empresa['cuit'])}
        client = wsfe.get_client(wsfe_wsdl)

        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]

        resultados = []
        ultimos    = {}

        for idx, row in df.iterrows():
            comp = row.to_dict()
            pv   = int(comp['punto_venta'])
            tipo = int(comp['tipo_cbte'])
            key  = (pv, tipo)

            try:
                if key not in ultimos:
                    ultimos[key] = wsfe.get_ultimo_comprobante(client, auth, pv, tipo)
                ultimos[key] += 1
                nro = ultimos[key]

                fecha_raw = comp['fecha']
                if hasattr(fecha_raw, 'strftime'):
                    fecha_str = fecha_raw.strftime('%Y%m%d')
                else:
                    fecha_str = datetime.strptime(str(fecha_raw)[:10], '%Y-%m-%d').strftime('%Y%m%d')
                comp['fecha'] = fecha_str

                result = wsfe.procesar_comprobante(
                    client, auth, empresa['cuit'], pv, tipo, comp, nro
                )
                det = result.FeDetResp.FECAEDetResponse[0]

                if det.Resultado == 'A':
                    resultados.append({
                        'fila': idx + 2, 'nro': nro, 'resultado': 'APROBADO',
                        'cae': _str_cae(det.CAE), 'vto_cae': str(det.CAEFchVto), 'obs': '',
                        'tipo_cbte': tipo, 'tipo_nombre': TIPO_NOMBRE.get(tipo, f'Tipo {tipo}'),
                        'tipo_grupo': _tipo_grupo(tipo),
                    })
                else:
                    obs = ''
                    if det.Observaciones:
                        obs = '; '.join(o.Msg for o in det.Observaciones.Obs)
                    resultados.append({
                        'fila': idx + 2, 'nro': nro, 'resultado': 'RECHAZADO',
                        'cae': '', 'vto_cae': '', 'obs': obs,
                        'tipo_cbte': tipo, 'tipo_nombre': TIPO_NOMBRE.get(tipo, f'Tipo {tipo}'),
                        'tipo_grupo': _tipo_grupo(tipo),
                    })

            except Exception as e:
                resultados.append({
                    'fila': idx + 2, 'nro': 0, 'resultado': 'ERROR',
                    'cae': '', 'vto_cae': '', 'obs': str(e),
                    'tipo_cbte': tipo, 'tipo_nombre': TIPO_NOMBRE.get(tipo, f'Tipo {tipo}'),
                    'tipo_grupo': _tipo_grupo(tipo),
                })

        _guardar_resultado(path, _resultado_path_actual(empresa_id), resultados)

        aprobados = sum(1 for r in resultados if r['resultado'] == 'APROBADO')
        return jsonify({
            'ok': True,
            'resultados': resultados,
            'resumen': {
                'total': len(resultados),
                'aprobados': aprobados,
                'rechazados': len(resultados) - aprobados,
            },
        })

    except Exception as e:
        print(f"\n=== ERROR /procesar ===\n{traceback.format_exc()}\n=====\n")
        return jsonify({'error': str(e)}), 500


def _guardar_resultado(src_path: str, dest_path: str, resultados: list):
    RES_HEADERS = ['Nro_Cbte', 'Resultado', 'CAE', 'Vto_CAE', 'Observaciones']
    verde    = PatternFill(fill_type='solid', fgColor='C6EFCE')
    rojo     = PatternFill(fill_type='solid', fgColor='FFC7CE')
    amarillo = PatternFill(fill_type='solid', fgColor='FFEB9C')

    wb_src = load_workbook(src_path)
    ws_src = wb_src.active
    n_src_cols  = ws_src.max_column
    src_headers = [ws_src.cell(1, c).value for c in range(1, n_src_cols + 1)]

    # Cargar destino existente o crear uno nuevo con encabezados
    if os.path.exists(dest_path):
        wb_dest = load_workbook(dest_path)
        ws_dest = wb_dest.active
        n_dest_cols  = ws_dest.max_column
        dest_headers = [ws_dest.cell(1, c).value for c in range(1, n_dest_cols + 1)]
    else:
        wb_dest = Workbook()
        ws_dest = wb_dest.active
        dest_headers = src_headers + RES_HEADERS
        for i, h in enumerate(dest_headers, 1):
            ws_dest.cell(1, i, h)

    # Mapa nombre_de_columna → índice en destino (case-insensitive)
    dest_col_map = {
        str(h).lower().strip(): i + 1
        for i, h in enumerate(dest_headers) if h is not None
    }

    result_map = {r['fila']: r for r in resultados}

    for row_idx in range(2, ws_src.max_row + 1):
        r    = result_map.get(row_idx, {})
        res  = r.get('resultado', '')
        fill = verde if res == 'APROBADO' else (rojo if res == 'RECHAZADO' else amarillo)

        new_row = ws_dest.max_row + 1

        # Escribir valores fuente mapeando por nombre de columna
        for ci, header in enumerate(src_headers, 1):
            val    = ws_src.cell(row_idx, ci).value
            h_key  = str(header).lower().strip() if header is not None else ''
            dest_c = dest_col_map.get(h_key, ci)
            ws_dest.cell(new_row, dest_c, val).fill = fill

        # Escribir columnas de resultado por nombre
        for val, hdr in [
            (r.get('nro', ''),                  'nro_cbte'),
            (res,                                'resultado'),
            (_str_cae(r.get('cae', '')),         'cae'),
            (_str_cae(r.get('vto_cae', '')),     'vto_cae'),
            (r.get('obs', ''),                   'observaciones'),
        ]:
            dest_c = dest_col_map.get(hdr)
            if dest_c:
                ws_dest.cell(new_row, dest_c, val).fill = fill

    wb_dest.save(dest_path)


@app.route('/api/resultados-mes')
@login_required
def api_resultados_mes():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    mes        = request.args.get('mes', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'resultados': [], 'mes': '', 'resumen': {'total': 0, 'aprobados': 0, 'rechazados': 0}})

    base_dir = os.path.join(UPLOAD, empresa['cuit'])

    # Si no se pide un mes específico, buscar el último con datos
    if not mes:
        if os.path.exists(base_dir):
            meses = sorted([
                d for d in os.listdir(base_dir)
                if os.path.isdir(os.path.join(base_dir, d))
                and re.match(r'\d{4}-\d{2}', d)
                and os.path.exists(os.path.join(base_dir, d, 'facturas_resultado.xlsx'))
            ], reverse=True)
            mes = meses[0] if meses else _mes_actual()
        else:
            mes = _mes_actual()

    path = os.path.join(base_dir, mes, 'facturas_resultado.xlsx')
    if not os.path.exists(path):
        return jsonify({'resultados': [], 'mes': mes, 'resumen': {'total': 0, 'aprobados': 0, 'rechazados': 0}})

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
        resultados = []
        for idx, row in df.iterrows():
            res = str(row.get('resultado', '')).upper()
            if res not in ('APROBADO', 'RECHAZADO', 'ERROR'):
                continue
            t     = int(row.get('tipo_cbte', 0))
            grupo = _tipo_grupo(t)
            email_raw = str(row.get('email', ''))
            email_val = email_raw.strip() if email_raw not in ('nan', 'None', '') else ''
            resultados.append({
                'fila':       int(idx) + 2,
                'nro':        int(row.get('nro_cbte', 0)) if res == 'APROBADO' else 0,
                'resultado':  'APROBADO' if res == 'APROBADO' else res,
                'cae':        _str_cae(row.get('cae', '')) if res == 'APROBADO' else '',
                'vto_cae':    _str_cae(row.get('vto_cae', '')) if res == 'APROBADO' else '',
                'obs':        str(row.get('observaciones', '')),
                'tipo_cbte':  t,
                'tipo_nombre': TIPO_NOMBRE.get(t, f'Tipo {t}') if t else '',
                'tipo_grupo': grupo,
                'email':      email_val,
            })
        aprobados = sum(1 for r in resultados if r['resultado'] == 'APROBADO')
        return jsonify({
            'mes':       mes,
            'resultados': resultados,
            'resumen': {
                'total':      len(resultados),
                'aprobados':  aprobados,
                'rechazados': len(resultados) - aprobados,
            },
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/meses-disponibles')
@login_required
def api_meses_disponibles():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'meses': []})
    base_dir = os.path.join(UPLOAD, empresa['cuit'])
    meses = []
    if os.path.exists(base_dir):
        meses = sorted([
            d for d in os.listdir(base_dir)
            if os.path.isdir(os.path.join(base_dir, d))
            and re.match(r'\d{4}-\d{2}', d)
            and os.path.exists(os.path.join(base_dir, d, 'facturas_resultado.xlsx'))
        ], reverse=True)
    return jsonify({'meses': meses})


@app.route('/api/consultar-cuit')
@login_required
def api_consultar_cuit():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    cuit_consulta = re.sub(r'[^0-9]', '', request.args.get('cuit', ''))

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403
    if not re.fullmatch(r'\d{11}', cuit_consulta):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400

    # Buscar primero en la base local (rápido, sin AFIP)
    cliente_local = _get_cliente(empresa_id, cuit_consulta)
    if cliente_local and cliente_local.get('domicilio'):
        return jsonify({
            'ok':          True,
            'razon_social': cliente_local.get('nombre', ''),
            'domicilio':    cliente_local['domicilio'],
            'estado':       cliente_local.get('estado', ''),
            'fuente':       'local',
        })

    cert_path = empresa.get('cert', '')
    key_path  = empresa.get('key', '')
    if not os.path.isfile(cert_path) or not os.path.isfile(key_path):
        return jsonify({'error': 'Esta empresa no tiene certificados configurados'}), 400

    try:
        import wspadron
    except ImportError:
        return jsonify({'error': 'Módulo wspadron no encontrado. Ejecutá actualizar.ps1 para obtenerlo.'}), 500

    wsaa_url    = _empresa_urls(empresa)[0]
    padron_wsdl = wspadron.PADRON_WSDL_HOMO if empresa.get('homologacion') else wspadron.PADRON_WSDL_PROD

    try:
        token, sign = wsaa.get_ticket('ws_sr_constancia_inscripcion', cert_path, key_path, wsaa_url, empresa['cuit'])
        data = wspadron.consultar_persona(token, sign, empresa['cuit'], cuit_consulta, padron_wsdl)

        # Actualizar datos en base local si ya existe el cliente
        clientes = _load_clientes(empresa_id)
        if cuit_consulta in clientes:
            clientes[cuit_consulta]['domicilio']     = data.get('domicilio', '')
            clientes[cuit_consulta]['estado']         = data.get('estado', '')
            clientes[cuit_consulta]['condicion_iva'] = data.get('condicion_iva', '')
            _save_clientes(empresa_id, clientes)

        return jsonify({'ok': True, **data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _emitir_comprobante_asociado(tipo_map, label, empresa_id, mes_orig, fila, user):
    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    result_path = _resultado_path(empresa_id, mes_orig)
    if not os.path.exists(result_path):
        return jsonify({'error': 'No se encontraron resultados para ese mes'}), 400

    df = pd.read_excel(result_path)
    df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
    idx = fila - 2
    if idx < 0 or idx >= len(df):
        return jsonify({'error': 'Fila no encontrada'}), 400

    orig      = df.iloc[idx].to_dict()
    tipo_orig = int(orig.get('tipo_cbte', 0))
    tipo_nuevo = tipo_map.get(tipo_orig)
    if not tipo_nuevo:
        return jsonify({'error': f'No se puede emitir {label} para tipo {tipo_orig}'}), 400

    cert_path = empresa.get('cert', '')
    key_path  = empresa.get('key', '')
    if not os.path.isfile(cert_path):
        return jsonify({'error': f'Certificado no encontrado: {cert_path}'}), 400
    if not os.path.isfile(key_path):
        return jsonify({'error': f'Clave no encontrada: {key_path}'}), 400

    wsaa_url, wsfe_wsdl = _empresa_urls(empresa)

    try:
        token, sign = wsaa.get_ticket('wsfe', cert_path, key_path, wsaa_url, empresa['cuit'])
        auth_data   = {'Token': token, 'Sign': sign, 'Cuit': int(empresa['cuit'])}
        client      = wsfe.get_client(wsfe_wsdl)

        pv         = int(orig.get('punto_venta', 0))
        ultimo     = wsfe.get_ultimo_comprobante(client, auth_data, pv, tipo_nuevo)
        nro        = ultimo + 1
        fecha_str  = datetime.today().strftime('%Y%m%d')

        comp = {c: orig.get(c, '') for c in COLUMNAS}
        comp['tipo_cbte'] = tipo_nuevo
        comp['fecha']     = fecha_str

        cbtes_asoc = [{'tipo': tipo_orig, 'pv': pv, 'nro': int(orig.get('nro_cbte', 0))}]

        result = wsfe.procesar_comprobante(
            client, auth_data, empresa['cuit'], pv, tipo_nuevo, comp, nro, cbtes_asoc=cbtes_asoc
        )
        det = result.FeDetResp.FECAEDetResponse[0]

        if det.Resultado != 'A':
            obs = '; '.join(o.Msg for o in det.Observaciones.Obs) if det.Observaciones else ''
            return jsonify({'error': f'AFIP rechazó el comprobante: {obs}'}), 400

        res_data = {'fila': 2, 'nro': nro, 'resultado': 'APROBADO',
                    'cae': _str_cae(det.CAE), 'vto_cae': str(det.CAEFchVto), 'obs': ''}

        import tempfile
        tmp = tempfile.mktemp(suffix='.xlsx')
        try:
            df_nuevo = pd.DataFrame([{c: comp.get(c, '') for c in COLUMNAS}])
            with pd.ExcelWriter(tmp, engine='openpyxl') as w:
                df_nuevo.to_excel(w, index=False)
            _guardar_resultado(tmp, _resultado_path_actual(empresa_id), [res_data])
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)

        return jsonify({
            'ok':      True,
            'tipo':    tipo_nuevo,
            'nombre':  TIPO_NOMBRE.get(tipo_nuevo, f'Tipo {tipo_nuevo}'),
            'nro':     nro,
            'cae':     _str_cae(det.CAE),
            'vto_cae': str(det.CAEFchVto),
        })

    except Exception as e:
        print(f"\n=== ERROR {label} ===\n{traceback.format_exc()}\n=====\n")
        return jsonify({'error': str(e)}), 500


@app.route('/api/registrar-manual', methods=['POST'])
@login_required
def api_registrar_manual():
    user       = _get_current_user()
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    mes        = (data.get('mes')        or _mes_actual()).strip()

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    cae     = (data.get('cae') or '').strip()
    nro_cbte = data.get('nro_cbte')
    if not cae:
        return jsonify({'error': 'El CAE es obligatorio'}), 400
    if not nro_cbte:
        return jsonify({'error': 'El número de comprobante es obligatorio'}), 400

    EXCEL_HEADERS = [
        'punto_venta', 'tipo_cbte', 'concepto', 'doc_tipo', 'doc_nro', 'razon_social',
        'fecha', 'imp_neto', 'alicuota', 'imp_iva', 'imp_total',
        'Nro_Cbte', 'Resultado', 'CAE', 'Vto_CAE', 'Observaciones',
    ]
    verde = PatternFill(fill_type='solid', fgColor='C6EFCE')

    dest_path = _resultado_path(empresa_id, mes)
    if os.path.exists(dest_path):
        wb = load_workbook(dest_path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        for i, h in enumerate(EXCEL_HEADERS, 1):
            ws.cell(1, i, h)

    fecha_raw = (data.get('fecha') or '').replace('-', '')     # YYYY-MM-DD → YYYYMMDD
    vto_raw   = (data.get('vto_cae') or '').replace('-', '')

    tipo_cbte = int(data.get('tipo_cbte', 1))
    row_vals  = [
        int(data.get('punto_venta', 1)),
        tipo_cbte,
        int(data.get('concepto', 2)),
        int(data.get('doc_tipo', 80)),
        (data.get('doc_nro') or '').strip(),
        (data.get('razon_social') or '').strip(),
        int(fecha_raw) if fecha_raw else '',
        float(data.get('imp_neto', 0) or 0),
        float(data.get('alicuota', 21) or 21),
        float(data.get('imp_iva', 0) or 0),
        float(data.get('imp_total', 0) or 0),
        int(nro_cbte),
        'APROBADO',
        cae,
        vto_raw,
        (data.get('observaciones') or '').strip(),
    ]

    new_row = ws.max_row + 1
    for col_idx, val in enumerate(row_vals, 1):
        ws.cell(new_row, col_idx, val).fill = verde
    wb.save(dest_path)

    return jsonify({
        'ok':          True,
        'fila':        new_row,
        'nro':         int(nro_cbte),
        'cae':         cae,
        'vto_cae':     vto_raw,
        'tipo_cbte':   tipo_cbte,
        'tipo_nombre': TIPO_NOMBRE.get(tipo_cbte, f'Tipo {tipo_cbte}'),
        'tipo_grupo':  _tipo_grupo(tipo_cbte),
        'mes':         mes,
    })


@app.route('/api/nota-credito', methods=['POST'])
@login_required
def api_nota_credito():
    user = _get_current_user()
    d    = request.get_json(force=True)
    return _emitir_comprobante_asociado(
        NC_TIPO_MAP, 'NC',
        d.get('empresa_id', '').strip(),
        d.get('mes', '').strip(),
        int(d.get('fila', 0)),
        user,
    )


@app.route('/api/nota-debito', methods=['POST'])
@login_required
def api_nota_debito():
    user = _get_current_user()
    d    = request.get_json(force=True)
    return _emitir_comprobante_asociado(
        ND_TIPO_MAP, 'ND',
        d.get('empresa_id', '').strip(),
        d.get('mes', '').strip(),
        int(d.get('fila', 0)),
        user,
    )


@app.route('/descargar')
@login_required
def descargar():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    mes        = request.args.get('mes', '').strip() or _mes_actual()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403
    path = _resultado_path(empresa_id, mes)
    if not os.path.exists(path):
        return 'No hay resultado disponible', 404
    nombre = f'facturas_{empresa["cuit"]}_{mes}_resultado.xlsx'
    return send_file(path, as_attachment=True, download_name=nombre)


@app.route('/descargar-pdfs')
@login_required
def descargar_pdfs():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    mes        = request.args.get('mes', '').strip() or _mes_actual()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403
    path = _resultado_path(empresa_id, mes)
    if not os.path.exists(path):
        return 'No hay resultados para este mes', 404

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
        buf_zip = io.BytesIO()
        count = 0
        with zipfile.ZipFile(buf_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
            for idx in range(len(df)):
                row = df.iloc[idx].to_dict()
                resultado_val = str(row.get('resultado', '')).upper()
                if resultado_val != 'APROBADO':
                    continue
                registro = {c: row.get(c, '') for c in COLUMNAS}
                registro['fecha'] = str(registro['fecha'])
                resultado = {
                    'nro':     int(row.get('nro_cbte', 0)),
                    'cae':     _str_cae(row.get('cae', '')),
                    'vto_cae': _str_cae(row.get('vto_cae', '')),
                }
                cliente    = _get_cliente(empresa_id, registro.get('doc_nro', ''))
                forma_pago = str(row.get('forma_pago', '') or '').strip()
                pdf_buf = factura_pdf.generar_pdf(empresa, registro, resultado,
                                                  cliente=cliente, forma_pago=forma_pago)
                pv  = int(registro['punto_venta'])
                nro = int(resultado['nro'])
                tipo = int(registro.get('tipo_cbte', 0))
                tipo_nombre = TIPO_NOMBRE.get(tipo, f'tipo{tipo}')
                tipo_nombre = tipo_nombre.replace(' ', '_')
                fname = f'{tipo_nombre}_{pv:04d}-{nro:08d}.pdf'
                zf.writestr(fname, pdf_buf.read())
                count += 1

        if count == 0:
            return 'No hay comprobantes aprobados para descargar', 404

        buf_zip.seek(0)
        nombre = f'comprobantes_{empresa["cuit"]}_{mes}.zip'
        return send_file(buf_zip, as_attachment=True, download_name=nombre,
                         mimetype='application/zip')
    except Exception as e:
        return f'Error al generar PDFs: {e}', 500


@app.route('/plantilla')
@login_required
def plantilla():
    df = pd.DataFrame([{
        'punto_venta': 6, 'tipo_cbte': 11, 'concepto': 2,
        'doc_tipo': 99, 'doc_nro': 0, 'razon_social': 'Consumidor Final',
        'fecha': datetime.today().strftime('%Y-%m-%d'),
        'imp_neto': 1000.00, 'alicuota': 0, 'imp_iva': 0.00, 'imp_total': 1000.00,
        'forma_pago': 'Contado',
        'email': '',
    }])
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Facturas')
    out.seek(0)
    return send_file(out, as_attachment=True, download_name='plantilla_facturas.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- PDF ---------------------------------------------------------------

@app.route('/pdf/<empresa_id>/<int:fila>')
@login_required
def pdf_desde_resultado(empresa_id, fila):
    user    = _get_current_user()
    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403

    mes  = request.args.get('mes', '').strip() or _mes_actual()
    path = _resultado_path(empresa_id, mes)
    if not os.path.exists(path):
        return 'No hay resultados guardados para esta empresa', 404

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
        idx = fila - 2
        if idx < 0 or idx >= len(df):
            return f'Fila {fila} no encontrada', 404

        row      = df.iloc[idx].to_dict()
        registro = {c: row.get(c, '') for c in COLUMNAS}
        registro['fecha'] = str(registro['fecha'])
        resultado = {
            'nro':     int(row.get('nro_cbte', 0)),
            'cae':     _str_cae(row.get('cae', '')),
            'vto_cae': _str_cae(row.get('vto_cae', '')),
        }
        cliente    = _get_cliente(empresa_id, registro.get('doc_nro', ''))
        forma_pago = str(row.get('forma_pago', '') or '').strip()

        pdf_buf = factura_pdf.generar_pdf(empresa, registro, resultado, cliente=cliente, forma_pago=forma_pago)
        pv      = int(registro['punto_venta'])
        nro     = int(resultado['nro'])
        return send_file(pdf_buf, as_attachment=False,
                         download_name=f'factura_{pv:04d}-{nro:08d}.pdf',
                         mimetype='application/pdf')
    except Exception as e:
        return f'Error al generar PDF: {e}', 500


@app.route('/imprimir', methods=['POST'])
@login_required
def imprimir():
    user       = _get_current_user()
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    registro  = data.get('registro')
    resultado = data.get('resultado')
    if not registro or not resultado:
        return jsonify({'error': 'Faltan datos del comprobante'}), 400

    try:
        cliente    = _get_cliente(empresa_id, registro.get('doc_nro', ''))
        forma_pago = str(registro.get('forma_pago', '') or '').strip()
        pdf_buf    = factura_pdf.generar_pdf(empresa, registro, resultado, cliente=cliente, forma_pago=forma_pago)
        pv         = int(registro.get('punto_venta', 0))
        nro        = int(resultado.get('nro', 0))
        return send_file(pdf_buf, as_attachment=False,
                         download_name=f'factura_{pv:04d}-{nro:08d}.pdf',
                         mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ---------- email -----------------------------------------------------------

EMAIL_CONFIG_PATH = os.path.join(BASE, 'email_config.json')

def _load_email_config() -> dict:
    if os.path.exists(EMAIL_CONFIG_PATH):
        with open(EMAIL_CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {'smtp_server': '', 'smtp_port': 587, 'smtp_user': '',
            'smtp_password': '', 'smtp_ssl': False, 'nombre_remitente': ''}

def _save_email_config(cfg: dict):
    with open(EMAIL_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


@app.route('/api/config-email', methods=['GET'])
@admin_required
def api_get_email_config():
    cfg  = _load_email_config()
    safe = dict(cfg)
    if safe.get('smtp_password'):
        safe['smtp_password'] = '••••••••'
    return jsonify(safe)


@app.route('/api/config-email', methods=['POST'])
@admin_required
def api_set_email_config():
    data = request.get_json(force=True)
    cfg  = _load_email_config()
    cfg['smtp_server']      = (data.get('smtp_server')      or '').strip()
    cfg['smtp_port']        = int(data.get('smtp_port')      or 587)
    cfg['smtp_user']        = (data.get('smtp_user')         or '').strip()
    cfg['nombre_remitente'] = (data.get('nombre_remitente')  or '').strip()
    cfg['smtp_ssl']         = bool(data.get('smtp_ssl', False))
    new_pw = (data.get('smtp_password') or '').strip()
    if new_pw and '•' not in new_pw:
        cfg['smtp_password'] = new_pw
    _save_email_config(cfg)
    return jsonify({'ok': True})


# ---------- AI config --------------------------------------------------------

AI_CONFIG_PATH = os.path.join(BASE, 'ai_config.json')

def _load_ai_config() -> dict:
    if os.path.exists(AI_CONFIG_PATH):
        with open(AI_CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {
        'providers': {
            'openai':    {'api_key': '', 'model': 'gpt-4o',                    'enabled': False},
            'anthropic': {'api_key': '', 'model': 'claude-sonnet-4-20250514', 'enabled': False},
            'google':    {'api_key': '', 'model': 'gemini-1.5-flash',          'enabled': False},
        },
        'default_provider': ''
    }

def _save_ai_config(cfg: dict):
    with open(AI_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


@app.route('/api/config-ai', methods=['GET'])
@admin_required
def api_get_ai_config():
    cfg  = _load_ai_config()
    safe = json.loads(json.dumps(cfg))          # deep copy
    for prov in safe.get('providers', {}).values():
        if prov.get('api_key'):
            prov['api_key'] = '••••••••'
    return jsonify(safe)


@app.route('/api/config-ai', methods=['POST'])
@admin_required
def api_set_ai_config():
    data = request.get_json(force=True)
    cfg  = _load_ai_config()

    incoming_providers = data.get('providers', {})
    for name in ('openai', 'anthropic', 'google'):
        inc  = incoming_providers.get(name, {})
        cur  = cfg['providers'].setdefault(name, {'api_key': '', 'model': '', 'enabled': False})
        new_key = inc.get('api_key')
        if new_key is not None:
            new_key = new_key.strip()
            if new_key == '':
                cur['api_key'] = ''
            elif '•' not in new_key:
                cur['api_key'] = new_key
        if inc.get('model'):
            cur['model'] = inc['model'].strip()
        cur['enabled'] = bool(inc.get('enabled', cur.get('enabled', False)))

    cfg['default_provider'] = (data.get('default_provider') or '').strip()
    _save_ai_config(cfg)
    return jsonify({'ok': True})


@app.route('/api/config-ai/test', methods=['POST'])
@admin_required
def api_test_ai():
    import requests as _req
    data = request.get_json(force=True)
    provider = (data.get('provider') or '').strip()
    cfg = _load_ai_config()
    prov = cfg.get('providers', {}).get(provider, {})
    api_key = prov.get('api_key', '')
    model = prov.get('model', '')
    if not api_key:
        return jsonify({'error': 'No hay API key configurada para este proveedor.'}), 400
    if not model:
        return jsonify({'error': 'No hay modelo configurado.'}), 400

    prompt = 'Respondé solamente con la palabra: OK'
    try:
        if provider == 'openai':
            resp = _req.post('https://api.openai.com/v1/chat/completions',
                headers={'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'},
                json={'model': model, 'messages': [{'role': 'user', 'content': prompt}], 'max_tokens': 10},
                timeout=30)
        elif provider == 'anthropic':
            resp = _req.post('https://api.anthropic.com/v1/messages',
                headers={'x-api-key': api_key, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json'},
                json={'model': model, 'max_tokens': 10, 'messages': [{'role': 'user', 'content': prompt}]},
                timeout=30)
        elif provider == 'google':
            resp = _req.post(
                f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}',
                headers={'Content-Type': 'application/json'},
                json={'contents': [{'parts': [{'text': prompt}]}]},
                timeout=30)
        else:
            return jsonify({'error': f'Proveedor desconocido: {provider}'}), 400

        if resp.status_code >= 400:
            detail = ''
            try:
                body = resp.json()
                err = body.get('error', {})
                if isinstance(err, dict):
                    detail = err.get('message', '')
                    etype = err.get('type', '')
                    if etype:
                        detail = f'[{etype}] {detail}'
                else:
                    detail = str(err)
            except Exception:
                detail = resp.text[:300]
            if api_key and api_key in detail:
                detail = detail.replace(api_key, '***')
            return jsonify({'error': f'Error {resp.status_code}: {detail or resp.reason}'}), 400

        return jsonify({'ok': True, 'message': f'Conexión exitosa con {provider} (modelo: {model})'})
    except _req.exceptions.ConnectionError:
        return jsonify({'error': 'No se pudo conectar. Verificá tu conexión a internet.'}), 500
    except _req.exceptions.Timeout:
        return jsonify({'error': 'Tiempo de espera agotado. Intentá de nuevo.'}), 500
    except Exception as e:
        msg = str(e)
        if api_key and api_key in msg:
            msg = msg.replace(api_key, '***')
        return jsonify({'error': f'Error: {msg}'}), 500


# ---------- PDF regex parser (sin IA) ----------------------------------------

def _parse_ib_pdf_regex(text: str, tipo: str) -> dict | None:
    import re
    if not text or len(text.strip()) < 30:
        return None

    def _buscar(patterns, txt):
        for p in patterns:
            m = re.search(p, txt, re.IGNORECASE | re.DOTALL)
            if m:
                return m.group(1).strip()
        return ''

    txt = text.replace('\r\n', '\n')

    if tipo == 'inscripcion':
        datos = {}
        datos['nro_inscripcion'] = _buscar([
            r'(?:n[°ºo]?\s*(?:de\s+)?inscripci[oó]n|inscripci[oó]n\s*n[°ºo]?)\s*[:\-]?\s*(\S+)',
            r'inscripci[oó]n[:\s]+(\d[\d\-/]+)',
        ], txt)
        datos['regimen'] = _buscar([
            r'[rr][eé]gimen\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        datos['estado'] = _buscar([
            r'[eE]stado\s*[:\-]?\s*(\w+)',
        ], txt)
        datos['actividad'] = _buscar([
            r'[aA]ctividad(?:\s+principal)?\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        datos['cod_actividad'] = _buscar([
            r'[cC][oó]d(?:igo)?\.?\s*(?:de\s+)?[aA]ctividad\s*[:\-]?\s*(\S+)',
        ], txt)
        datos['domicilio'] = _buscar([
            r'[dD]omicilio(?:\s+fiscal)?\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        datos['vigencia_desde'] = _buscar([
            r'[vV]igencia\s*(?:desde)?\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
            r'[dD]esde\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['vigencia_hasta'] = _buscar([
            r'[vV]igencia\s*(?:.*?)[hH]asta\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
            r'[hH]asta\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['periodo_desde'] = _buscar([
            r'[pP]er[ií]odo\s*(?:desde)?\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['nro_constancia'] = _buscar([
            r'[cC]onstancia\s*(?:n[°ºo]?)?\s*[:\-]?\s*(\d[\d\-]+)',
            r'[nN][°ºo]?\s*[cC]onstancia\s*[:\-]?\s*(\d[\d\-]+)',
        ], txt)
        datos['provincia'] = _buscar([
            r'[pP]rovincia\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        datos['organismo'] = _buscar([
            r'(?:API|ARBA|AGIP|DGR|ATER|ATP|DGIP|DPR)\b',
        ], txt)
        if datos['organismo']:
            datos['organismo'] = datos['organismo'].upper()
        datos['fecha_inicio_actividad'] = _buscar([
            r'[iI]nicio\s*(?:de\s+)?[aA]ctividad(?:es)?\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['datos_actualizados'] = _buscar([
            r'[dD]atos\s+[aA]ctualizados?\s*(?:al)?\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['categoria'] = _buscar([
            r'[cC]ategor[ií]a\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        filled = sum(1 for v in datos.values() if v)
        if filled < 2:
            return None
        return datos

    else:  # exencion
        datos = {}
        datos['nro_cuenta'] = _buscar([
            r'[cC]uenta\s*(?:n[°ºo]?)?\s*[:\-]?\s*(\S+)',
            r'[nN][°ºo]?\s*[cC]uenta\s*[:\-]?\s*(\S+)',
        ], txt)
        datos['actividad'] = _buscar([
            r'[aA]ctividad(?:\s+principal)?\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        datos['cod_actividad'] = _buscar([
            r'[cC][oó]d(?:igo)?\.?\s*(?:de\s+)?[aA]ctividad\s*[:\-]?\s*(\S+)',
        ], txt)
        datos['encuadre'] = _buscar([
            r'[eE]ncuadre(?:\s+normativo)?\s*[:\-]?\s*(.+?)(?:\n|$)',
            r'[aA]rt[ií]culo\s*.+?(?:\n|$)',
        ], txt)
        datos['nro_constancia'] = _buscar([
            r'[cC]onstancia\s*(?:n[°ºo]?)?\s*[:\-]?\s*(\d[\d\-]+)',
            r'[nN][°ºo]?\s*[cC]onstancia\s*[:\-]?\s*(\d[\d\-]+)',
        ], txt)
        datos['valida_hasta'] = _buscar([
            r'[vV][aá]lida?\s*[hH]asta\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
            r'[hH]asta\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
            r'[vV]encimiento\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['fecha_tramite'] = _buscar([
            r'[fF]echa\s*(?:de\s+)?[tT]r[aá]mite\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['fecha_inicio_actividad'] = _buscar([
            r'[iI]nicio\s*(?:de\s+)?[aA]ctividad(?:es)?\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        ], txt)
        datos['provincia'] = _buscar([
            r'[pP]rovincia\s*[:\-]?\s*(.+?)(?:\n|$)',
        ], txt)
        datos['organismo'] = _buscar([
            r'(?:API|ARBA|AGIP|DGR|ATER|ATP|DGIP|DPR)\b',
        ], txt)
        if datos.get('organismo'):
            datos['organismo'] = datos['organismo'].upper()
        filled = sum(1 for v in datos.values() if v)
        if filled < 2:
            return None
        return datos


# ---------- AI helpers -------------------------------------------------------

def _extract_pdf_text(file_bytes: bytes) -> str:
    text = ''
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            text += (page.extract_text() or '') + '\n'
    except (ImportError, Exception):
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    text += (page.extract_text() or '') + '\n'
        except (ImportError, Exception):
            pass
    return text.strip()


def _ai_post_with_retry(req_mod, url, max_retries=3, **kwargs):
    import time
    for attempt in range(max_retries + 1):
        resp = req_mod.post(url, **kwargs)
        if resp.status_code == 429 and attempt < max_retries:
            wait = (attempt + 1) * 30
            time.sleep(wait)
            continue
        if resp.status_code >= 400:
            detail = ''
            try:
                body = resp.json()
                err = body.get('error', {})
                if isinstance(err, dict):
                    detail = err.get('message', '')
                else:
                    detail = str(err)
            except Exception:
                detail = resp.text[:300]
            raise ValueError(f'Error {resp.status_code}: {detail or resp.reason}')
        return resp
    raise ValueError(f'Error {resp.status_code} después de {max_retries} reintentos')


def _call_ai_extract(text: str, tipo: str, file_bytes: bytes = None) -> dict:
    import requests as _req
    import base64

    cfg = _load_ai_config()
    provider_name = cfg.get('default_provider', '')
    if not provider_name:
        raise ValueError('No hay proveedor de IA configurado como predeterminado.')

    prov = cfg.get('providers', {}).get(provider_name)
    if not prov or not prov.get('enabled'):
        raise ValueError(f'El proveedor "{provider_name}" no está habilitado.')
    api_key = prov.get('api_key', '')
    if not api_key:
        raise ValueError(f'El proveedor "{provider_name}" no tiene API key configurada.')
    model = prov.get('model', '')

    if tipo == 'inscripcion':
        campos = ('nro_inscripcion, regimen, categoria, estado, periodo_desde, '
                  'domicilio, cod_actividad, actividad, fecha_inicio_actividad, '
                  'nro_constancia, datos_actualizados, vigencia_desde, vigencia_hasta, '
                  'organismo, provincia')
        tipo_label = 'inscripción'
    else:
        campos = ('nro_cuenta, cod_actividad, actividad, fecha_inicio_actividad, '
                  'encuadre, nro_constancia, fecha_tramite, valida_hasta, '
                  'organismo, provincia')
        tipo_label = 'exención'

    instruccion = (
        f'Extraé los datos de esta constancia de {tipo_label} de Ingresos Brutos de Argentina. '
        f'Devolvé SOLO un JSON con estos campos exactos: {campos}. '
        'Sin explicaciones, solo el JSON.'
    )

    use_vision = not text and file_bytes is not None
    pdf_b64 = base64.b64encode(file_bytes).decode() if use_vision else None

    if provider_name == 'openai':
        if use_vision:
            raise ValueError(
                'OpenAI no soporta PDFs directamente. '
                'Instalá PyPDF2 (pip install PyPDF2) o cambiá el proveedor '
                'a Google (Gemini) o Anthropic (Claude) en Admin → Inteligencia Artificial.'
            )
        content = instruccion + f'\n\nTexto del documento:\n---\n{text}\n---'
        resp = _ai_post_with_retry(
            _req,
            'https://api.openai.com/v1/chat/completions',
            headers={'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'},
            json={
                'model': model,
                'messages': [{'role': 'user', 'content': content}],
                'response_format': {'type': 'json_object'},
            },
            timeout=90,
        )
        result = resp.json()['choices'][0]['message']['content']

    elif provider_name == 'anthropic':
        if use_vision:
            content = [
                {'type': 'document', 'source': {
                    'type': 'base64',
                    'media_type': 'application/pdf',
                    'data': pdf_b64,
                }},
                {'type': 'text', 'text': instruccion},
            ]
        else:
            content = instruccion + f'\n\nTexto del documento:\n---\n{text}\n---'
        resp = _ai_post_with_retry(
            _req,
            'https://api.anthropic.com/v1/messages',
            headers={
                'x-api-key': api_key,
                'anthropic-version': '2023-06-01',
                'anthropic-beta': 'pdfs-2024-09-25',
                'Content-Type': 'application/json',
            },
            json={
                'model': model,
                'max_tokens': 2048,
                'messages': [{'role': 'user', 'content': content}],
            },
            timeout=90,
        )
        result = resp.json()['content'][0]['text']

    elif provider_name == 'google':
        if use_vision:
            parts = [
                {'inline_data': {'mime_type': 'application/pdf', 'data': pdf_b64}},
                {'text': instruccion},
            ]
        else:
            parts = [{'text': instruccion + f'\n\nTexto del documento:\n---\n{text}\n---'}]
        resp = _ai_post_with_retry(
            _req,
            f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}',
            headers={'Content-Type': 'application/json'},
            json={
                'contents': [{'parts': parts}],
                'generationConfig': {'responseMimeType': 'application/json'},
            },
            timeout=90,
        )
        result = resp.json()['candidates'][0]['content']['parts'][0]['text']

    else:
        raise ValueError(f'Proveedor desconocido: {provider_name}')

    result = result.strip()
    if result.startswith('```'):
        result = result.split('\n', 1)[-1].rsplit('```', 1)[0].strip()
    return json.loads(result)


# ---------- IB PDF upload + AI extraction ------------------------------------

@app.route('/api/clientes/ib-pdf', methods=['POST'])
@login_required
def api_clientes_ib_pdf():
    user = _get_current_user()
    empresa_id = request.form.get('empresa_id', '').strip()
    cuit       = request.form.get('cuit', '').strip()
    tipo       = request.form.get('tipo', '').strip()          # inscripcion | exencion

    if not empresa_id or not cuit or tipo not in ('inscripcion', 'exencion'):
        return jsonify({'error': 'Faltan parámetros (empresa_id, cuit, tipo).'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Sin acceso a esta empresa.'}), 403

    archivo = request.files.get('archivo')
    if not archivo or not archivo.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Se requiere un archivo PDF.'}), 400

    file_bytes = archivo.read()
    text = _extract_pdf_text(file_bytes)

    datos = None
    if text:
        datos = _parse_ib_pdf_regex(text, tipo)

    if not datos:
        try:
            datos = _call_ai_extract(text, tipo, file_bytes=file_bytes)
        except Exception as e:
            msg = str(e)
            cfg = _load_ai_config()
            for prov in cfg.get('providers', {}).values():
                k = prov.get('api_key', '')
                if k and k in msg:
                    msg = msg.replace(k, '***')
            if '429' in msg:
                msg = 'Demasiadas peticiones al proveedor de IA. Esperá 1 minuto y volvé a intentar.'
            if not text:
                msg = 'No se pudo extraer texto del PDF. Verificá que el PDF tenga texto seleccionable (no escaneo/imagen).'
            return jsonify({'error': msg}), 500

    # Guardar en el cliente
    clientes = _load_clientes(empresa_id)
    cliente  = clientes.get(cuit)
    if not cliente:
        return jsonify({'error': f'Cliente con CUIT {cuit} no encontrado.'}), 404

    ib = cliente.setdefault('ingresos_brutos', {})
    ib[tipo] = datos
    clientes[cuit] = cliente
    _save_clientes(empresa_id, clientes)

    return jsonify({'ok': True, 'datos': datos})


@app.route('/api/enviar-factura', methods=['POST'])
@login_required
def api_enviar_factura():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders as _enc

    user       = _get_current_user()
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    fila       = int(data.get('fila', 0))
    mes        = (data.get('mes') or _mes_actual()).strip()
    email_dst  = (data.get('email_destino') or '').strip()
    asunto     = (data.get('asunto') or '').strip()
    mensaje    = (data.get('mensaje') or '').strip()

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403
    if not email_dst:
        return jsonify({'error': 'Email del destinatario es requerido'}), 400

    # Prioridad: config de la empresa; si no tiene, usa la config global
    if empresa.get('smtp_server') and empresa.get('smtp_user') and empresa.get('smtp_password'):
        cfg = {
            'smtp_server':      empresa['smtp_server'],
            'smtp_port':        empresa.get('smtp_port', 587),
            'smtp_user':        empresa['smtp_user'],
            'smtp_password':    empresa['smtp_password'],
            'smtp_ssl':         empresa.get('smtp_ssl', False),
            'nombre_remitente': empresa.get('nombre_remitente', ''),
        }
    else:
        cfg = _load_email_config()
        if not cfg.get('smtp_server') or not cfg.get('smtp_user') or not cfg.get('smtp_password'):
            return jsonify({'error': 'Correo no configurado. Configuralo en Admin → Empresas o en Admin → Correo (global).'}), 400

    path = _resultado_path(empresa_id, mes)
    if not os.path.exists(path):
        return jsonify({'error': 'No hay resultados guardados para ese mes'}), 400

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
        idx = fila - 2
        if idx < 0 or idx >= len(df):
            return jsonify({'error': 'Fila no encontrada'}), 400

        row      = df.iloc[idx].to_dict()
        registro = {c: row.get(c, '') for c in COLUMNAS}
        registro['fecha'] = str(registro['fecha'])
        resultado = {
            'nro':     int(row.get('nro_cbte', 0)),
            'cae':     _str_cae(row.get('cae', '')),
            'vto_cae': _str_cae(row.get('vto_cae', '')),
        }
        cliente    = _get_cliente(empresa_id, registro.get('doc_nro', ''))
        forma_pago = str(row.get('forma_pago', '') or '').strip()
        pdf_buf    = factura_pdf.generar_pdf(empresa, registro, resultado, cliente=cliente, forma_pago=forma_pago)
        pv         = int(registro['punto_venta'])
        nro        = int(resultado['nro'])
        filename   = f'factura_{pv:05d}-{nro:08d}.pdf'
    except Exception as e:
        return jsonify({'error': f'Error generando PDF: {e}'}), 500

    if not asunto:
        tipo_nom = factura_pdf.TIPO_NOMBRE.get(int(registro.get('tipo_cbte', 0)), 'Comprobante')
        asunto   = f'{tipo_nom} Nro {pv:05d}-{nro:08d} - {empresa["nombre"]}'
    if not mensaje:
        mensaje = (
            f'Estimado/a {row.get("razon_social", "")}:\n\n'
            f'Adjuntamos el comprobante electronico.\n\n'
            f'Saludos,\n{cfg.get("nombre_remitente") or empresa["nombre"]}'
        )

    msg            = MIMEMultipart()
    remitente      = f'{cfg.get("nombre_remitente") or empresa["nombre"]} <{cfg["smtp_user"]}>'
    msg['From']    = remitente
    msg['To']      = email_dst
    msg['Subject'] = asunto
    msg.attach(MIMEText(mensaje, 'plain', 'utf-8'))

    part = MIMEBase('application', 'pdf')
    part.set_payload(pdf_buf.read())
    _enc.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(part)

    try:
        port = int(cfg.get('smtp_port', 587))
        srv  = cfg['smtp_server']
        usr  = cfg['smtp_user']
        pw   = cfg['smtp_password']

        if cfg.get('smtp_ssl') or port == 465:
            with smtplib.SMTP_SSL(srv, port, timeout=20) as s:
                s.login(usr, pw)
                s.sendmail(usr, email_dst, msg.as_bytes())
        else:
            with smtplib.SMTP(srv, port, timeout=20) as s:
                s.ehlo(); s.starttls(); s.ehlo()
                s.login(usr, pw)
                s.sendmail(usr, email_dst, msg.as_bytes())

        return jsonify({'ok': True})
    except smtplib.SMTPAuthenticationError:
        return jsonify({'error': 'Error de autenticacion SMTP. Verifica usuario y contrasena.'}), 500
    except Exception as e:
        return jsonify({'error': f'Error al enviar: {e}'}), 500


@app.route('/api/clientes')
@login_required
def api_clientes():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403
    clientes = _load_clientes(empresa_id)
    lista = sorted(clientes.values(), key=lambda x: x.get('nombre', ''))
    return jsonify({'clientes': lista, 'total': len(lista)})


@app.route('/api/clientes/importar', methods=['POST'])
@admin_required
def api_clientes_importar():
    empresa_id = (request.form.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

    f = request.files.get('archivo')
    if not f or not f.filename:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    try:
        df = pd.read_excel(f)
        df.columns = [c.strip().upper() for c in df.columns]

        cuit_col   = next((c for c in df.columns if 'CUIT' in c), None)
        nombre_col = next((c for c in df.columns if 'CLIENTE' in c or 'NOMBRE' in c or 'RAZON' in c), None)
        codigo_col = next((c for c in df.columns if c in ('Nº', 'NRO', 'CODIGO', 'CÓDIGO', 'COD', 'N°', 'Nº'.upper())), None)

        if not cuit_col or not nombre_col:
            return jsonify({'error': f'No se encontraron columnas CUIT y CLIENTE. Columnas: {list(df.columns)}'}), 400

        clientes   = _load_clientes(empresa_id)
        importados = 0
        actualizados = 0

        for _, row in df.iterrows():
            raw_cuit   = row.get(cuit_col)
            raw_nombre = row.get(nombre_col)
            if not raw_cuit or pd.isna(raw_cuit):
                continue
            cuit_str = str(int(float(raw_cuit))).strip()
            if not cuit_str or cuit_str == '0':
                continue
            nombre = str(raw_nombre or '').strip()
            codigo = ''
            if codigo_col:
                raw_cod = row.get(codigo_col)
                if raw_cod is not None and not pd.isna(raw_cod):
                    codigo = str(int(float(raw_cod))) if isinstance(raw_cod, (int, float)) else str(raw_cod).strip()

            if cuit_str in clientes:
                changed = False
                if clientes[cuit_str].get('nombre') != nombre:
                    clientes[cuit_str]['nombre'] = nombre
                    changed = True
                if codigo and clientes[cuit_str].get('codigo') != codigo:
                    clientes[cuit_str]['codigo'] = codigo
                    changed = True
                if changed:
                    actualizados += 1
            else:
                clientes[cuit_str] = {
                    'cuit':      cuit_str,
                    'nombre':    nombre,
                    'codigo':    codigo,
                    'domicilio': '',
                    'estado':    '',
                }
                importados += 1

        _save_clientes(empresa_id, clientes)
        return jsonify({'ok': True, 'importados': importados,
                        'actualizados': actualizados, 'total': len(clientes)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/clientes/sync', methods=['POST'])
@admin_required
def api_clientes_sync():
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

    cert_path = empresa.get('cert', '')
    key_path  = empresa.get('key', '')
    if not os.path.isfile(cert_path) or not os.path.isfile(key_path):
        return jsonify({'error': 'Esta empresa no tiene certificados configurados'}), 400

    try:
        import wspadron
    except ImportError:
        return jsonify({'error': 'Módulo wspadron no encontrado. Ejecutá actualizar.ps1.'}), 500

    clientes = _load_clientes(empresa_id)
    if not clientes:
        return jsonify({'error': 'No hay clientes importados para esta empresa'}), 400

    wsaa_url    = _empresa_urls(empresa)[0]
    padron_wsdl = wspadron.PADRON_WSDL_HOMO if empresa.get('homologacion') else wspadron.PADRON_WSDL_PROD

    try:
        token, sign = wsaa.get_ticket('ws_sr_constancia_inscripcion', cert_path, key_path, wsaa_url, empresa['cuit'])
    except Exception as e:
        return jsonify({'error': f'Error obteniendo ticket AFIP: {e}'}), 500

    actualizados = 0
    errores      = []

    for cuit, cliente in clientes.items():
        try:
            info = wspadron.consultar_persona(token, sign, empresa['cuit'], cuit, padron_wsdl)
            if info.get('domicilio'):
                cliente['domicilio'] = info['domicilio']
            if info.get('estado'):
                cliente['estado']    = info['estado']
            if info.get('condicion_iva'):
                cliente['condicion_iva'] = info['condicion_iva']
            actualizados += 1
        except Exception as e:
            errores.append(f'{cuit}: {str(e)[:60]}')

    _save_clientes(empresa_id, clientes)
    return jsonify({'ok': True, 'actualizados': actualizados,
                    'errores': errores[:10], 'total': len(clientes)})


@app.route('/api/clientes/ventas')
@login_required
def api_clientes_ventas():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    cuit       = request.args.get('cuit', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403
    if not cuit:
        return jsonify({'error': 'CUIT requerido'}), 400

    registros = _leer_registros_reporte(empresa_id, desde='', hasta='', cliente='')
    ventas = [r for r in registros if r.get('doc_nro') == cuit]
    ventas.sort(key=lambda r: r.get('fecha', ''), reverse=True)
    return jsonify({'ventas': ventas, 'total': len(ventas)})


@app.route('/api/clientes/compras')
@login_required
def api_clientes_compras():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    cuit       = request.args.get('cuit', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    compras = _load_compras(empresa_id)
    if cuit:
        compras = [c for c in compras if c.get('cuit_emisor') == cuit]
    compras.sort(key=lambda c: c.get('fecha', ''), reverse=True)
    return jsonify({'compras': compras, 'total': len(compras)})


@app.route('/api/clientes/foto', methods=['POST'])
@login_required
def api_clientes_foto_upload():
    user       = _get_current_user()
    empresa_id = (request.form.get('empresa_id') or '').strip()
    cuit       = (request.form.get('cuit') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403
    if not cuit:
        return jsonify({'error': 'CUIT requerido'}), 400

    f = request.files.get('foto')
    if not f or not f.filename:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.jpg', '.jpeg', '.png', '.webp'):
        return jsonify({'error': 'Solo se permiten imágenes JPG, PNG o WEBP'}), 400

    carpeta = os.path.join(FOTOS_CLI, empresa_id)
    os.makedirs(carpeta, exist_ok=True)
    dest = os.path.join(carpeta, f'{cuit}{ext}')

    for old_ext in ('.jpg', '.jpeg', '.png', '.webp'):
        old = os.path.join(carpeta, f'{cuit}{old_ext}')
        if os.path.exists(old):
            os.remove(old)

    f.save(dest)

    clientes = _load_clientes(empresa_id)
    if cuit in clientes:
        clientes[cuit]['foto'] = f'/static/clientes/{empresa_id}/{cuit}{ext}'
        _save_clientes(empresa_id, clientes)

    return jsonify({'ok': True, 'foto_url': f'/static/clientes/{empresa_id}/{cuit}{ext}'})


@app.route('/api/clientes/importar-ib', methods=['POST'])
@admin_required
def api_clientes_importar_ib():
    empresa_id = (request.form.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

    f = request.files.get('archivo')
    if not f or not f.filename:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    try:
        wb = load_workbook(f)
        clientes = _load_clientes(empresa_id)
        actualizados = 0

        for ws in wb.worksheets:
            title_val = str(ws.cell(1, 1).value or '').upper()
            headers = [str(ws.cell(2, c).value or '').strip() for c in range(1, ws.max_column + 1)]

            cuit_idx = next((i for i, h in enumerate(headers) if 'CUIT' in h.upper()), None)
            if cuit_idx is None:
                continue

            is_exencion = 'EXEN' in title_val

            for r in range(3, ws.max_row + 1):
                raw_cuit = ws.cell(r, cuit_idx + 1).value
                if not raw_cuit:
                    continue
                cuit_str = str(raw_cuit).replace('-', '').strip()
                if not cuit_str or cuit_str == 'None':
                    continue

                if cuit_str not in clientes:
                    nombre_idx = next((i for i, h in enumerate(headers) if 'Nombre' in h or 'Razón' in h or 'Razon' in h or 'NOMBRE' in h.upper()), None)
                    nombre = str(ws.cell(r, nombre_idx + 1).value or '').strip() if nombre_idx is not None else ''
                    clientes[cuit_str] = {
                        'cuit': cuit_str, 'nombre': nombre,
                        'codigo': '', 'domicilio': '', 'estado': '',
                    }

                ib = clientes[cuit_str].get('ingresos_brutos', {})

                if is_exencion:
                    ex = {}
                    for i, h in enumerate(headers):
                        val = ws.cell(r, i + 1).value
                        if val is None:
                            continue
                        hl = h.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
                        if 'cuenta' in hl:
                            ex['nro_cuenta'] = str(val).strip()
                        elif 'actividad' in hl and 'descripci' in hl:
                            ex['actividad'] = str(val).strip()
                        elif 'actividad' in hl and 'codigo' in hl:
                            ex['cod_actividad'] = str(val).strip()
                        elif 'actividad' in hl and 'fecha' in hl:
                            ex['fecha_inicio_actividad'] = str(val).strip()
                        elif 'encuadre' in hl or 'normativo' in hl:
                            ex['encuadre'] = str(val).strip()
                        elif 'valida' in hl or ('vigencia' in hl and 'hasta' in hl):
                            ex['valida_hasta'] = str(val).strip()
                        elif 'constancia' in hl and h.startswith('N'):
                            ex['nro_constancia'] = str(val).strip()
                        elif 'tramite' in hl:
                            ex['fecha_tramite'] = str(val).strip()
                        elif 'organismo' in hl:
                            ex['organismo'] = str(val).strip()
                        elif 'provincia' in hl:
                            ex['provincia'] = str(val).strip()
                    ib['exencion'] = ex
                else:
                    insc = {}
                    for i, h in enumerate(headers):
                        val = ws.cell(r, i + 1).value
                        if val is None:
                            continue
                        hl = h.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
                        if 'inscripcion' in hl:
                            insc['nro_inscripcion'] = str(val).strip()
                        elif 'regimen' in hl:
                            insc['regimen'] = str(val).strip()
                        elif 'categori' in hl:
                            insc['categoria'] = str(val).strip()
                        elif 'estado' in hl:
                            insc['estado'] = str(val).strip()
                        elif 'periodo' in hl and 'desde' in hl:
                            insc['periodo_desde'] = str(val).strip()
                        elif 'domicilio' in hl:
                            insc['domicilio'] = str(val).strip()
                        elif 'actividad' in hl and 'descripcion' in hl:
                            insc['actividad'] = str(val).strip()
                        elif 'actividad' in hl and 'codigo' in hl:
                            insc['cod_actividad'] = str(val).strip()
                        elif 'actividad' in hl and 'fecha' in hl:
                            insc['fecha_inicio_actividad'] = str(val).strip()
                        elif 'constancia' in hl:
                            insc['nro_constancia'] = str(val).strip()
                        elif 'actualizado' in hl:
                            insc['datos_actualizados'] = str(val).strip()
                        elif 'vigencia' in hl and 'desde' in hl:
                            insc['vigencia_desde'] = str(val).strip()
                        elif 'vigencia' in hl and 'hasta' in hl:
                            insc['vigencia_hasta'] = str(val).strip()
                        elif 'organismo' in hl:
                            insc['organismo'] = str(val).strip()
                        elif 'provincia' in hl:
                            insc['provincia'] = str(val).strip()
                    ib['inscripcion'] = insc

                clientes[cuit_str]['ingresos_brutos'] = ib
                actualizados += 1

        _save_clientes(empresa_id, clientes)
        return jsonify({'ok': True, 'actualizados': actualizados, 'total': len(clientes)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/clientes/importar-compras', methods=['POST'])
@login_required
def api_clientes_importar_compras():
    user = _get_current_user()
    empresa_id = (request.form.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'No tenés acceso a esta empresa'}), 403

    f = request.files.get('archivo')
    if not f or not f.filename:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    # Validar CUIT del archivo contra la empresa seleccionada
    cuit_empresa = empresa['cuit'].replace('-', '').strip()
    nombre_archivo = f.filename or ''
    m_cuit = re.search(r'(\d{11})', nombre_archivo)
    if m_cuit and m_cuit.group(1) != cuit_empresa:
        return jsonify({
            'error': f'El CUIT del archivo ({m_cuit.group(1)}) no coincide con la empresa seleccionada ({cuit_empresa} - {empresa["nombre"]}). '
                     f'Verificá que estás importando el archivo correcto para esta empresa.'
        }), 400

    try:
        wb = load_workbook(f)
        ws = wb.active

        fila_headers = 1
        v1 = str(ws.cell(1, 1).value or '').strip()
        v2 = str(ws.cell(2, 1).value or '').strip()
        if v1 and v2 and 'Fecha' in v2:
            fila_headers = 2

        headers = [str(ws.cell(fila_headers, c).value or '').strip()
                   for c in range(1, ws.max_column + 1)]

        compras   = _load_compras(empresa_id)
        caes_exist = {c.get('cae') for c in compras if c.get('cae')}

        importados = 0
        duplicados = 0
        errores    = []

        for row_idx in range(fila_headers + 1, ws.max_row + 1):
            row = {headers[i]: ws.cell(row_idx, i + 1).value
                   for i in range(len(headers))}

            fecha_raw = row.get('Fecha')
            if not fecha_raw:
                continue

            try:
                if hasattr(fecha_raw, 'strftime'):
                    fecha_dt = fecha_raw
                else:
                    fecha_dt = datetime.strptime(str(fecha_raw).strip(), '%d/%m/%Y')

                cae = _str_cae(row.get('Cód. Autorización', '') or row.get('CAE', ''))
                if cae and cae in caes_exist:
                    duplicados += 1
                    continue

                cuit_emisor_raw = row.get('Nro. Doc. Emisor') or row.get('CUIT Emisor') or 0
                cuit_emisor = str(int(float(cuit_emisor_raw))).strip() if cuit_emisor_raw else ''
                denominacion = str(row.get('Denominación Emisor') or '').strip()

                tipo_str   = str(row.get('Tipo') or '').strip()
                imp_total  = round(float(row.get('Imp. Total') or 0), 2)
                imp_neto_raw = row.get('Imp. Neto Gravado') or row.get('Neto Gravado Total')
                imp_iva_raw  = row.get('IVA') or row.get('Total IVA')
                imp_neto   = round(float(imp_neto_raw or 0), 2)
                imp_iva    = round(float(imp_iva_raw or 0), 2)

                punto_venta = int(float(row.get('Punto de Venta') or 0))
                nro_desde   = int(float(row.get('Número Desde') or 0))

                compras.append({
                    'fecha':       fecha_dt.strftime('%Y-%m-%d'),
                    'tipo':        tipo_str,
                    'punto_venta': punto_venta,
                    'nro_desde':   nro_desde,
                    'cuit_emisor': cuit_emisor,
                    'denominacion_emisor': denominacion,
                    'imp_total':   imp_total,
                    'imp_neto':    imp_neto,
                    'imp_iva':     imp_iva,
                    'cae':         cae,
                    'moneda':      str(row.get('Moneda') or 'PES').strip(),
                })

                if cae:
                    caes_exist.add(cae)
                importados += 1

            except Exception as ex:
                errores.append(f'Fila {row_idx}: {ex}')

        _save_compras(empresa_id, compras)
        return jsonify({
            'ok': True, 'importados': importados,
            'duplicados': duplicados, 'total': len(compras),
            'errores': errores[:10],
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/enviar-todos', methods=['POST'])
@login_required
def api_enviar_todos():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders as _enc

    user       = _get_current_user()
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    mes        = (data.get('mes') or _mes_actual()).strip()

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    if empresa.get('smtp_server') and empresa.get('smtp_user') and empresa.get('smtp_password'):
        cfg = {
            'smtp_server':      empresa['smtp_server'],
            'smtp_port':        empresa.get('smtp_port', 587),
            'smtp_user':        empresa['smtp_user'],
            'smtp_password':    empresa['smtp_password'],
            'smtp_ssl':         empresa.get('smtp_ssl', False),
            'nombre_remitente': empresa.get('nombre_remitente', ''),
        }
    else:
        cfg = _load_email_config()
        if not cfg.get('smtp_server') or not cfg.get('smtp_user') or not cfg.get('smtp_password'):
            return jsonify({'error': 'Correo no configurado. Configuralo en Admin → Empresas o en Admin → Correo (global).'}), 400

    path = _resultado_path(empresa_id, mes)
    if not os.path.exists(path):
        return jsonify({'error': 'No hay resultados guardados para ese mes'}), 400

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
    except Exception as e:
        return jsonify({'error': f'Error leyendo resultados: {e}'}), 500

    enviados  = []
    errores   = []
    omitidos  = 0

    port = int(cfg.get('smtp_port', 587))
    srv  = cfg['smtp_server']
    usr  = cfg['smtp_user']
    pw   = cfg['smtp_password']

    try:
        if cfg.get('smtp_ssl') or port == 465:
            smtp_conn = smtplib.SMTP_SSL(srv, port, timeout=30)
        else:
            smtp_conn = smtplib.SMTP(srv, port, timeout=30)
            smtp_conn.ehlo(); smtp_conn.starttls(); smtp_conn.ehlo()
        smtp_conn.login(usr, pw)
    except smtplib.SMTPAuthenticationError:
        return jsonify({'error': 'Error de autenticacion SMTP. Verifica usuario y contrasena.'}), 500
    except Exception as e:
        return jsonify({'error': f'No se pudo conectar al servidor de correo: {e}'}), 500

    try:
        for idx, row in df.iterrows():
            res = str(row.get('resultado', '')).upper()
            if res != 'APROBADO':
                omitidos += 1
                continue

            email_raw = str(row.get('email', ''))
            email_dst = email_raw.strip() if email_raw not in ('nan', 'None', '') else ''
            if not email_dst:
                omitidos += 1
                continue

            fila = int(idx) + 2
            try:
                registro = {c: row.get(c, '') for c in COLUMNAS}
                registro['fecha'] = str(registro['fecha'])
                resultado = {
                    'nro':     int(row.get('nro_cbte', 0)),
                    'cae':     _str_cae(row.get('cae', '')),
                    'vto_cae': _str_cae(row.get('vto_cae', '')),
                }
                cliente    = _get_cliente(empresa_id, registro.get('doc_nro', ''))
                forma_pago = str(row.get('forma_pago', '') or '').strip()
                pdf_buf    = factura_pdf.generar_pdf(empresa, registro, resultado, cliente=cliente, forma_pago=forma_pago)
                pv       = int(registro['punto_venta'])
                nro      = int(resultado['nro'])
                filename = f'factura_{pv:05d}-{nro:08d}.pdf'

                tipo_nom = TIPO_NOMBRE.get(int(registro.get('tipo_cbte', 0)), 'Comprobante')
                asunto   = f'{tipo_nom} Nro {pv:05d}-{nro:08d} - {empresa["nombre"]}'
                remitente_nom = cfg.get('nombre_remitente') or empresa['nombre']
                mensaje  = (
                    f'Estimado/a {row.get("razon_social", "")}:\n\n'
                    f'Adjuntamos el comprobante electrónico.\n\n'
                    f'Saludos,\n{remitente_nom}'
                )

                msg            = MIMEMultipart()
                msg['From']    = f'{remitente_nom} <{usr}>'
                msg['To']      = email_dst
                msg['Subject'] = asunto
                msg.attach(MIMEText(mensaje, 'plain', 'utf-8'))

                part = MIMEBase('application', 'pdf')
                part.set_payload(pdf_buf.read())
                _enc.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=filename)
                msg.attach(part)

                smtp_conn.sendmail(usr, email_dst, msg.as_bytes())
                enviados.append({'fila': fila, 'email': email_dst, 'nro': nro})
            except Exception as e:
                errores.append({'fila': fila, 'email': email_dst, 'error': str(e)})
    finally:
        try:
            smtp_conn.quit()
        except Exception:
            pass

    return jsonify({
        'ok':      True,
        'enviados': len(enviados),
        'errores':  errores,
        'omitidos': omitidos,
    })


# ---------- generador de CSR -------------------------------------------------

@app.route('/generar-csr', methods=['POST'])
@admin_required
def generar_csr():
    data   = request.get_json(force=True)
    cuit   = (data.get('cuit')   or '').strip()
    nombre = (data.get('nombre') or '').strip()
    email  = (data.get('email')  or '').strip()

    if not cuit or not nombre:
        return jsonify({'error': 'CUIT y nombre son obligatorios'}), 400
    if not re.fullmatch(r'\d{11}', cuit):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400

    cuit_dir = os.path.join(CERTS, cuit)
    os.makedirs(cuit_dir, exist_ok=True)
    key_path = os.path.join(cuit_dir, f'{cuit}_clave.key')
    csr_path = os.path.join(cuit_dir, f'{cuit}.csr')

    try:
        openssl = encontrar_openssl()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    try:
        r1 = subprocess.run([openssl, 'genrsa', '-out', key_path, '2048'],
                            capture_output=True)
        if r1.returncode != 0:
            raise Exception(r1.stderr.decode('utf-8', errors='replace'))

        subject = f'/C=AR/O={nombre}/serialNumber=CUIT {cuit}/CN={cuit}'
        if email:
            subject += f'/emailAddress={email}'

        r2 = subprocess.run([openssl, 'req', '-new', '-key', key_path,
                             '-out', csr_path, '-subj', subject],
                            capture_output=True)
        if r2.returncode != 0:
            raise Exception(r2.stderr.decode('utf-8', errors='replace'))

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(key_path, f'{cuit}_clave.key')
            zf.write(csr_path, f'{cuit}.csr')
        buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name=f'certificado_afip_{cuit}.zip',
                         mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ---------- reportes ---------------------------------------------------------

_TIPOS_FACTURA = {1, 6, 11, 51, 201, 206, 211}
_TIPOS_NC      = {3, 8, 13, 53, 203, 208, 213}
_TIPOS_ND      = {2, 7, 12, 52, 202, 207, 212}

def _tipo_grupo(tipo_cbte: int) -> str:
    if tipo_cbte in _TIPOS_NC:      return 'nc'
    if tipo_cbte in _TIPOS_ND:      return 'nd'
    if tipo_cbte in _TIPOS_FACTURA: return 'factura'
    return 'factura'

def _leer_registros_reporte(empresa_id: str, desde: str, hasta: str,
                             cliente: str, tipo: str = '') -> list[dict]:
    """Lee todos los resultados de una empresa filtrando por mes, cliente y tipo."""
    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return []
    base_dir = os.path.join(UPLOAD, empresa['cuit'])
    if not os.path.exists(base_dir):
        return []

    meses = sorted([
        d for d in os.listdir(base_dir)
        if os.path.isdir(os.path.join(base_dir, d)) and re.match(r'\d{4}-\d{2}', d)
    ])
    if desde:
        meses = [m for m in meses if m >= desde]
    if hasta:
        meses = [m for m in meses if m <= hasta]

    registros = []
    caes_vistos = set()
    for mes in meses:
        path = os.path.join(base_dir, mes, 'facturas_resultado.xlsx')
        if not os.path.exists(path):
            continue
        try:
            df = pd.read_excel(path)
            df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
            for idx, row in df.iterrows():
                if str(row.get('resultado', '')).upper() != 'APROBADO':
                    continue
                cae = _str_cae(row.get('cae', ''))
                if cae and cae in caes_vistos:
                    continue
                if cae:
                    caes_vistos.add(cae)
                razon = str(row.get('razon_social', ''))
                if cliente and cliente.lower() not in razon.lower():
                    continue
                t = int(row.get('tipo_cbte', 0))
                grupo = _tipo_grupo(t)
                if tipo and grupo != tipo:
                    continue
                registros.append({
                    'mes':          mes,
                    'fila':         int(idx) + 2,
                    'empresa_id':   empresa_id,
                    'fecha':        str(row.get('fecha', ''))[:10],
                    'razon_social': razon,
                    'doc_nro':      str(row.get('doc_nro', '')).split('.')[0].strip(),
                    'punto_venta':  int(row.get('punto_venta', 0)),
                    'nro_cbte':     int(row.get('nro_cbte', 0)),
                    'tipo_cbte':    t,
                    'tipo_nombre':  TIPO_NOMBRE.get(t, f'Tipo {t}'),
                    'tipo_grupo':   grupo,
                    'imp_neto':     float(row.get('imp_neto', 0)),
                    'imp_iva':      float(row.get('imp_iva', 0)),
                    'imp_total':    float(row.get('imp_total', 0)),
                    'cae':          cae,
                    'vto_cae':      _str_cae(row.get('vto_cae', '')),
                })
        except Exception as e:
            print(f"Error leyendo {path}: {e}")
    return registros


@app.route('/api/reportes/diagnostico')
@login_required
def api_reportes_diagnostico():
    """Devuelve info de diagnóstico: comprobantes por mes, duplicados por CAE y por cbte."""
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa or not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    base_dir = os.path.join(UPLOAD, empresa['cuit'])
    if not os.path.exists(base_dir):
        return jsonify({'error': 'Sin datos'}), 404

    meses = sorted([
        d for d in os.listdir(base_dir)
        if os.path.isdir(os.path.join(base_dir, d)) and re.match(r'\d{4}-\d{2}', d)
    ])

    all_caes = {}
    all_cbtes = {}
    por_mes = {}
    for mes in meses:
        path = os.path.join(base_dir, mes, 'facturas_resultado.xlsx')
        if not os.path.exists(path):
            continue
        try:
            df = pd.read_excel(path)
            df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]
            count_mes = 0
            for idx, row in df.iterrows():
                if str(row.get('resultado', '')).upper() != 'APROBADO':
                    continue
                count_mes += 1
                cae = _str_cae(row.get('cae', ''))
                t   = int(row.get('tipo_cbte', 0))
                pv  = int(row.get('punto_venta', 0))
                nro = int(row.get('nro_cbte', 0))
                cbte_key = f"{t}-{pv}-{nro}"
                if cae:
                    all_caes.setdefault(cae, []).append(mes)
                all_cbtes.setdefault(cbte_key, []).append(mes)
            por_mes[mes] = count_mes
        except Exception:
            pass

    dup_caes  = {k: v for k, v in all_caes.items() if len(v) > 1}
    dup_cbtes = {k: v for k, v in all_cbtes.items() if len(v) > 1}

    return jsonify({
        'empresa': empresa['nombre'],
        'cuit': empresa['cuit'],
        'por_mes': por_mes,
        'total_aprobados': sum(por_mes.values()),
        'caes_unicos': len(all_caes),
        'cbtes_unicos': len(all_cbtes),
        'duplicados_cae': dup_caes,
        'cantidad_dup_cae': len(dup_caes),
        'duplicados_cbte': dup_cbtes,
        'cantidad_dup_cbte': len(dup_cbtes),
    })


@app.route('/reportes')
@login_required
def reportes():
    user     = _get_current_user()
    empresas = _user_empresas(user)
    # Meses disponibles: desde el más antiguo hasta hoy
    mes_hasta = _mes_actual()
    mes_desde = (datetime.today().replace(day=1)
                 .replace(month=1)).strftime('%Y-%m')  # enero de este año
    return render_template('reportes.html', empresas=empresas,
                           current_user=user,
                           mes_desde=mes_desde, mes_hasta=mes_hasta)


@app.route('/api/reportes')
@login_required
def api_reportes():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    desde      = request.args.get('desde', '').strip()
    hasta      = request.args.get('hasta', '').strip()
    cliente    = request.args.get('cliente', '').strip()
    tipo       = request.args.get('tipo', '').strip()

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    registros = _leer_registros_reporte(empresa_id, desde, hasta, cliente, tipo)

    # Totales por grupo
    sum_fc = sum(r['imp_total'] for r in registros if r['tipo_grupo'] == 'factura')
    sum_nc = sum(r['imp_total'] for r in registros if r['tipo_grupo'] == 'nc')
    sum_nd = sum(r['imp_total'] for r in registros if r['tipo_grupo'] == 'nd')
    cant_fc = sum(1 for r in registros if r['tipo_grupo'] == 'factura')
    cant_nc = sum(1 for r in registros if r['tipo_grupo'] == 'nc')
    cant_nd = sum(1 for r in registros if r['tipo_grupo'] == 'nd')
    monto_neto = round(sum_fc - sum_nc + sum_nd, 2)

    # Totales por mes (usando el neto: facturas - NC + ND)
    por_mes: dict[str, float] = {}
    for r in registros:
        signo = -1 if r['tipo_grupo'] == 'nc' else 1
        por_mes[r['mes']] = round(por_mes.get(r['mes'], 0) + r['imp_total'] * signo, 2)

    return jsonify({
        'registros': registros,
        'totales': {
            'cantidad':    cant_fc + cant_nd,
            'cantidad_nc': cant_nc,
            'cantidad_total': len(registros),
            'monto':       round(sum_fc + sum_nd, 2),
            'monto_nc':    round(sum_nc, 2),
            'monto_neto':  monto_neto,
        },
        'por_mes': por_mes,
    })


@app.route('/api/reportes/exportar')
@login_required
def api_reportes_exportar():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    desde      = request.args.get('desde', '').strip()
    hasta      = request.args.get('hasta', '').strip()
    cliente    = request.args.get('cliente', '').strip()
    tipo       = request.args.get('tipo', '').strip()

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403

    registros = _leer_registros_reporte(empresa_id, desde, hasta, cliente, tipo)
    if not registros:
        return 'No hay datos para exportar', 404

    df = pd.DataFrame(registros, columns=[
        'mes', 'fecha', 'razon_social', 'punto_venta',
        'nro_cbte', 'tipo_nombre', 'tipo_grupo', 'imp_neto', 'imp_iva', 'imp_total', 'cae'
    ])
    df.columns = ['Mes', 'Fecha', 'Cliente', 'Pto. Venta',
                  'Nro. Cbte', 'Tipo Cbte', '_grupo', 'Neto', 'IVA', 'Total', 'CAE']

    # Calcular neto (Facturas + ND - NC)
    mask_nc = df['_grupo'] == 'nc'
    neto_total = round(df.loc[~mask_nc, 'Total'].sum() - df.loc[mask_nc, 'Total'].sum(), 2)
    neto_neto  = round(df.loc[~mask_nc, 'Neto'].sum()  - df.loc[mask_nc, 'Neto'].sum(), 2)
    neto_iva   = round(df.loc[~mask_nc, 'IVA'].sum()   - df.loc[mask_nc, 'IVA'].sum(), 2)
    cant_fc_nd = int((~mask_nc).sum())
    df.drop(columns=['_grupo'], inplace=True)

    # Fila de totales
    total_row = pd.DataFrame([{
        'Mes': '', 'Fecha': '', 'Cliente': 'NETO FACTURADO',
        'Pto. Venta': '', 'Nro. Cbte': cant_fc_nd, 'Tipo Cbte': '',
        'Neto': neto_neto, 'IVA': neto_iva,
        'Total': neto_total, 'CAE': '',
    }])
    df = pd.concat([df, total_row], ignore_index=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Reporte')
        ws = w.sheets['Reporte']
        # Ancho de columnas
        for col in ws.columns:
            max_len = max(len(str(c.value or '')) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
    out.seek(0)

    nombre = f"reporte_{empresa['cuit']}_{desde or 'inicio'}_{hasta or _mes_actual()}.xlsx"
    return send_file(out, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- backup / restore -------------------------------------------------

EMPRESAS_FILE = os.path.join(BASE, 'empresas.json')
USUARIOS_FILE = os.path.join(BASE, 'usuarios.json')

# --- Mapeos para importación desde AFIP --------------------------------------
_AFIP_DOC_TIPO = {
    'CUIT': 80, 'CUIL': 86, 'CDI': 87, 'LE': 91, 'LC': 90,
    'DNI': 96, 'PASAPORTE': 94, 'CI EXTRANJERA': 89,
    'SIN IDENTIFICAR': 99, '': 99,
}

def _afip_doc_tipo(s: str) -> int:
    return _AFIP_DOC_TIPO.get(str(s or '').strip().upper(), 99)

def _afip_tipo_cbte(s: str) -> int:
    m = re.match(r'^\s*(\d+)', str(s or ''))
    return int(m.group(1)) if m else 0

def _afip_alicuota(row: dict) -> float:
    for col, rate in [('IVA 27%', 27), ('IVA 21%', 21),
                      ('IVA 10,5%', 10.5), ('IVA 5%', 5), ('IVA 2,5%', 2.5)]:
        try:
            if row.get(col) and float(row[col]) > 0:
                return rate
        except (ValueError, TypeError):
            pass
    return 0


@app.route('/api/importar-afip', methods=['POST'])
@login_required
def api_importar_afip():
    user = _get_current_user()
    empresa_id = (request.form.get('empresa_id') or '').strip()
    empresa    = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'No tenés acceso a esta empresa'}), 403

    f = request.files.get('archivo')
    if not f or not f.filename:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    # Validar CUIT del archivo contra la empresa seleccionada
    cuit_empresa = empresa['cuit'].replace('-', '').strip()
    nombre_archivo = f.filename or ''
    m_cuit = re.search(r'(\d{11})', nombre_archivo)
    if m_cuit and m_cuit.group(1) != cuit_empresa:
        return jsonify({
            'error': f'El CUIT del archivo ({m_cuit.group(1)}) no coincide con la empresa seleccionada ({cuit_empresa} - {empresa["nombre"]}). '
                     f'Verificá que estás importando el archivo correcto para esta empresa.'
        }), 400

    EXCEL_HEADERS = COLUMNAS + ['Nro_Cbte', 'Resultado', 'CAE', 'Vto_CAE', 'Observaciones']

    try:
        wb_src = load_workbook(f)
        ws_src = wb_src.active

        # Detectar fila de encabezados: fila 1 o fila 2 (AFIP pone título en fila 1)
        fila_headers = 1
        if ws_src.cell(1, 1).value and ws_src.cell(2, 1).value == 'Fecha':
            fila_headers = 2

        headers = [str(ws_src.cell(fila_headers, c).value or '').strip()
                   for c in range(1, ws_src.max_column + 1)]

        importados = 0
        duplicados = 0
        errores    = []
        meses_ok   = set()

        for row_idx in range(fila_headers + 1, ws_src.max_row + 1):
            row = {headers[i]: ws_src.cell(row_idx, i + 1).value
                   for i in range(len(headers))}

            fecha_raw = row.get('Fecha')
            if not fecha_raw:
                continue

            try:
                # Fecha: puede venir como datetime o string DD/MM/YYYY
                if hasattr(fecha_raw, 'strftime'):
                    fecha_dt = fecha_raw
                else:
                    fecha_dt = datetime.strptime(str(fecha_raw).strip(), '%d/%m/%Y')

                mes      = fecha_dt.strftime('%Y-%m')
                fecha_str = fecha_dt.strftime('%Y-%m-%d')
                vto_cae  = (fecha_dt + timedelta(days=10)).strftime('%Y-%m-%d')

                tipo_cbte  = _afip_tipo_cbte(row.get('Tipo', ''))
                punto_venta = int(float(row.get('Punto de Venta') or 1))
                nro_cbte   = int(float(row.get('Número Desde') or 0))
                cae        = _str_cae(row.get('Cód. Autorización', ''))
                doc_tipo   = _afip_doc_tipo(row.get('Tipo Doc. Receptor', ''))
                doc_nro_raw = row.get('Nro. Doc. Receptor', 0)
                doc_nro    = str(int(float(doc_nro_raw or 0))) if doc_nro_raw else '0'
                razon_social = str(row.get('Denominación Receptor') or '').strip()

                imp_total  = round(float(row.get('Imp. Total') or 0), 2)
                imp_iva    = round(float(row.get('Total IVA') or 0), 2)
                neto_grav  = row.get('Neto Gravado Total')
                imp_neto   = round(float(neto_grav), 2) if neto_grav else round(imp_total - imp_iva, 2)
                alicuota   = _afip_alicuota(row)

                # Ruta del resultado para este mes
                dest_path = _resultado_path(empresa_id, mes)
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)

                # Verificar duplicado por CAE
                if cae and os.path.exists(dest_path):
                    wb_chk = load_workbook(dest_path)
                    ws_chk = wb_chk.active
                    hdrs_chk = [str(ws_chk.cell(1, c).value or '').strip().lower()
                                for c in range(1, ws_chk.max_column + 1)]
                    try:
                        cae_col = hdrs_chk.index('cae') + 1
                        for cr in range(2, ws_chk.max_row + 1):
                            if _str_cae(ws_chk.cell(cr, cae_col).value) == cae:
                                duplicados += 1
                                raise StopIteration
                    except StopIteration:
                        continue

                # Crear o cargar el destino
                if os.path.exists(dest_path):
                    wb_dest = load_workbook(dest_path)
                    ws_dest = wb_dest.active
                else:
                    wb_dest = Workbook()
                    ws_dest = wb_dest.active
                    for ci, h in enumerate(EXCEL_HEADERS, 1):
                        ws_dest.cell(1, ci, h)

                verde = PatternFill(fill_type='solid', fgColor='C6EFCE')
                vals  = [
                    punto_venta, tipo_cbte, 1, doc_tipo, doc_nro,
                    razon_social, fecha_str, imp_neto, alicuota, imp_iva, imp_total,
                    nro_cbte, 'APROBADO', cae, vto_cae, 'Importado de AFIP',
                ]
                new_row = ws_dest.max_row + 1
                for ci, v in enumerate(vals, 1):
                    ws_dest.cell(new_row, ci, v).fill = verde

                wb_dest.save(dest_path)
                importados += 1
                meses_ok.add(mes)

            except StopIteration:
                pass
            except Exception as ex:
                errores.append(f'Fila {row_idx}: {ex}')

        return jsonify({
            'ok':        True,
            'importados': importados,
            'duplicados': duplicados,
            'errores':    errores[:10],
            'meses':      sorted(meses_ok),
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/eliminar-por-pv', methods=['POST'])
@admin_required
def api_eliminar_por_pv():
    """Elimina comprobantes de puntos de venta específicos de una empresa."""
    data = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    pvs_raw = data.get('puntos_venta', [])

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

    pvs = set()
    for p in pvs_raw:
        try:
            pvs.add(int(p))
        except (ValueError, TypeError):
            pass
    if not pvs:
        return jsonify({'error': 'Indicá al menos un punto de venta'}), 400

    base_dir = os.path.join(UPLOAD, empresa['cuit'])
    if not os.path.exists(base_dir):
        return jsonify({'error': 'Sin datos para esta empresa'}), 404

    meses = sorted([
        d for d in os.listdir(base_dir)
        if os.path.isdir(os.path.join(base_dir, d)) and re.match(r'\d{4}-\d{2}', d)
    ])

    total_eliminados = 0
    meses_tocados = []

    for mes in meses:
        path = os.path.join(base_dir, mes, 'facturas_resultado.xlsx')
        if not os.path.exists(path):
            continue
        try:
            wb = load_workbook(path)
            ws = wb.active
            hdrs = [str(ws.cell(1, c).value or '').strip().lower().replace(' ', '_')
                    for c in range(1, ws.max_column + 1)]
            try:
                pv_col = hdrs.index('punto_venta') + 1
            except ValueError:
                continue

            filas_eliminar = []
            for r in range(2, ws.max_row + 1):
                try:
                    pv_val = int(float(ws.cell(r, pv_col).value or 0))
                except (ValueError, TypeError):
                    continue
                if pv_val in pvs:
                    filas_eliminar.append(r)

            if not filas_eliminar:
                continue

            for r in reversed(filas_eliminar):
                ws.delete_rows(r)

            wb.save(path)
            total_eliminados += len(filas_eliminar)
            meses_tocados.append(mes)
        except Exception as e:
            print(f"Error limpiando {path}: {e}")

    return jsonify({
        'ok': True,
        'eliminados': total_eliminados,
        'puntos_venta': sorted(pvs),
        'meses': meses_tocados,
        'empresa': empresa['nombre'],
    })


@app.route('/api/check-update')
@login_required
def api_check_update():
    import urllib.request
    from datetime import timedelta

    now = datetime.now()
    # Consultar GitHub como máximo una vez por día
    if _update_cache['checked_at'] and (now - _update_cache['checked_at']) < timedelta(hours=24):
        latest = _update_cache['latest']
    else:
        try:
            with urllib.request.urlopen(_UPDATE_URL, timeout=4) as r:
                latest = r.read().decode().strip()
            _update_cache['latest']     = latest
            _update_cache['checked_at'] = now
        except Exception:
            latest = None

    def _ver_tuple(v):
        try:
            return tuple(int(x) for x in v.split('.'))
        except Exception:
            return (0,)

    is_newer = False
    if latest:
        is_newer = _ver_tuple(latest) > _ver_tuple(APP_VERSION)

    return jsonify({
        'current': APP_VERSION,
        'latest':  latest,
        'update':  is_newer,
    })


@app.route('/admin/backup')
@admin_required
def admin_backup():
    """Genera un ZIP con uploads/, empresas.json y usuarios.json."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Datos de configuración
        for fname in ('empresas.json', 'usuarios.json'):
            fpath = os.path.join(BASE, fname)
            if os.path.exists(fpath):
                zf.write(fpath, fname)
        # Facturas (uploads/)
        for root, dirs, files in os.walk(UPLOAD):
            for file in files:
                filepath = os.path.join(root, file)
                arcname  = os.path.relpath(filepath, BASE)
                zf.write(filepath, arcname)
    buf.seek(0)
    fecha = datetime.today().strftime('%Y%m%d_%H%M')
    return send_file(buf, as_attachment=True,
                     download_name=f'arca_backup_{fecha}.zip',
                     mimetype='application/zip')


@app.route('/admin/restore', methods=['POST'])
@admin_required
def admin_restore():
    """Restaura un backup ZIP generado por /admin/backup."""
    f = request.files.get('backup')
    if not f or not f.filename.endswith('.zip'):
        return jsonify({'error': 'Seleccioná un archivo .zip de backup válido'}), 400

    try:
        with zipfile.ZipFile(f, 'r') as zf:
            for name in zf.namelist():
                # Seguridad: solo restaurar uploads/ y los JSON permitidos
                if name.startswith('uploads/') or name in ('empresas.json', 'usuarios.json'):
                    # Evitar path traversal
                    dest = os.path.normpath(os.path.join(BASE, name))
                    if not dest.startswith(BASE):
                        continue
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                    with zf.open(name) as src, open(dest, 'wb') as dst:
                        dst.write(src.read())
        return jsonify({'ok': True, 'msg': 'Backup restaurado correctamente. Reiniciá el servidor.'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ---------- clientes (página independiente) ------------------------------------

@app.route('/clientes')
@login_required
def clientes():
    user     = _get_current_user()
    empresas = _user_empresas(user)
    return render_template('clientes.html', empresas=empresas,
                           current_user=user)


# ---------- administración ---------------------------------------------------

@app.route('/admin')
@admin_required
def admin():
    user     = _get_current_user()
    empresas = EmpresaRepository.get_all()
    usuarios = UsuarioRepository.get_all_public()
    return render_template('admin.html', empresas=empresas, usuarios=usuarios,
                           certs_dir=CERTS, current_user=user)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
