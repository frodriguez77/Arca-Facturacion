import functools
import io
import os
import re
import subprocess
import traceback
import uuid
import zipfile
from datetime import datetime

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

BASE   = os.path.dirname(os.path.abspath(__file__))
UPLOAD = os.path.join(BASE, 'uploads')
CERTS  = os.path.join(BASE, 'certificados')

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

os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(CERTS,  exist_ok=True)


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

    if not nombre or not cuit or not cert or not key:
        return jsonify({'error': 'Nombre, CUIT, certificado y clave son obligatorios'}), 400
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

    if not nombre or not cuit or not cert or not key:
        return jsonify({'error': 'Nombre, CUIT, certificado y clave son obligatorios'}), 400
    if not re.fullmatch(r'\d{11}', cuit):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400
    if EmpresaRepository.cuit_exists(cuit, exclude_id=empresa_id):
        return jsonify({'error': f'Ya existe otra empresa con CUIT {cuit}'}), 400

    EmpresaRepository.update(empresa_id, {
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
    })
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

        registros = []
        for _, row in df.iterrows():
            r = {}
            for c in COLUMNAS:
                v = row[c]
                if hasattr(v, 'strftime'):
                    v = v.strftime('%Y-%m-%d')
                r[c] = str(v) if v is not None else ''
            registros.append(r)

        return jsonify({'ok': True, 'registros': registros, 'total': len(registros)})
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
                        'cae': det.CAE, 'vto_cae': str(det.CAEFchVto), 'obs': '',
                    })
                else:
                    obs = ''
                    if det.Observaciones:
                        obs = '; '.join(o.Msg for o in det.Observaciones.Obs)
                    resultados.append({
                        'fila': idx + 2, 'nro': nro, 'resultado': 'RECHAZADO',
                        'cae': '', 'vto_cae': '', 'obs': obs,
                    })

            except Exception as e:
                resultados.append({
                    'fila': idx + 2, 'nro': 0, 'resultado': 'ERROR',
                    'cae': '', 'vto_cae': '', 'obs': str(e),
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
    n_src_cols = ws_src.max_column

    # Cargar destino existente o crear uno nuevo con encabezados
    if os.path.exists(dest_path):
        wb_dest = load_workbook(dest_path)
        ws_dest = wb_dest.active
    else:
        wb_dest = Workbook()
        ws_dest = wb_dest.active
        src_headers = [ws_src.cell(1, c).value for c in range(1, n_src_cols + 1)]
        for i, h in enumerate(src_headers + RES_HEADERS, 1):
            ws_dest.cell(1, i, h)

    result_map = {r['fila']: r for r in resultados}

    for row_idx in range(2, ws_src.max_row + 1):
        r    = result_map.get(row_idx, {})
        res  = r.get('resultado', '')
        fill = verde if res == 'APROBADO' else (rojo if res == 'RECHAZADO' else amarillo)

        src_vals = [ws_src.cell(row_idx, c).value for c in range(1, n_src_cols + 1)]
        res_vals = [r.get('nro', ''), res, r.get('cae', ''), r.get('vto_cae', ''), r.get('obs', '')]

        new_row = ws_dest.max_row + 1
        for col_idx, val in enumerate(src_vals + res_vals, 1):
            ws_dest.cell(new_row, col_idx, val).fill = fill

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
            resultados.append({
                'fila':      int(idx) + 2,
                'nro':       int(row.get('nro_cbte', 0)) if res == 'APROBADO' else 0,
                'resultado': 'APROBADO' if res == 'APROBADO' else res,
                'cae':       str(row.get('cae', '')) if res == 'APROBADO' else '',
                'vto_cae':   str(row.get('vto_cae', '')) if res == 'APROBADO' else '',
                'obs':       str(row.get('observaciones', '')),
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
                    'cae': det.CAE, 'vto_cae': str(det.CAEFchVto), 'obs': ''}

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
            'cae':     det.CAE,
            'vto_cae': str(det.CAEFchVto),
        })

    except Exception as e:
        print(f"\n=== ERROR {label} ===\n{traceback.format_exc()}\n=====\n")
        return jsonify({'error': str(e)}), 500


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


@app.route('/plantilla')
@login_required
def plantilla():
    df = pd.DataFrame([{
        'punto_venta': 6, 'tipo_cbte': 11, 'concepto': 2,
        'doc_tipo': 99, 'doc_nro': 0, 'razon_social': 'Consumidor Final',
        'fecha': datetime.today().strftime('%Y-%m-%d'),
        'imp_neto': 1000.00, 'alicuota': 0, 'imp_iva': 0.00, 'imp_total': 1000.00,
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
            'cae':     str(row.get('cae', '')),
            'vto_cae': str(row.get('vto_cae', '')),
        }

        pdf_buf = factura_pdf.generar_pdf(empresa, registro, resultado)
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
        pdf_buf = factura_pdf.generar_pdf(empresa, registro, resultado)
        pv      = int(registro.get('punto_venta', 0))
        nro     = int(resultado.get('nro', 0))
        return send_file(pdf_buf, as_attachment=False,
                         download_name=f'factura_{pv:04d}-{nro:08d}.pdf',
                         mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


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

def _leer_registros_reporte(empresa_id: str, desde: str, hasta: str, cliente: str) -> list[dict]:
    """Lee todos los resultados de una empresa filtrando por mes y cliente."""
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
                razon = str(row.get('razon_social', ''))
                if cliente and cliente.lower() not in razon.lower():
                    continue
                registros.append({
                    'mes':          mes,
                    'fila':         int(idx) + 2,
                    'empresa_id':   empresa_id,
                    'fecha':        str(row.get('fecha', ''))[:10],
                    'razon_social': razon,
                    'punto_venta':  int(row.get('punto_venta', 0)),
                    'nro_cbte':     int(row.get('nro_cbte', 0)),
                    'tipo_cbte':    int(row.get('tipo_cbte', 0)),
                    'imp_neto':     float(row.get('imp_neto', 0)),
                    'imp_iva':      float(row.get('imp_iva', 0)),
                    'imp_total':    float(row.get('imp_total', 0)),
                    'cae':          str(row.get('cae', '')),
                })
        except Exception as e:
            print(f"Error leyendo {path}: {e}")
    return registros


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

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'Acceso denegado'}), 403

    registros = _leer_registros_reporte(empresa_id, desde, hasta, cliente)

    # Totales
    monto_total = sum(r['imp_total'] for r in registros)
    por_mes: dict[str, float] = {}
    for r in registros:
        por_mes[r['mes']] = round(por_mes.get(r['mes'], 0) + r['imp_total'], 2)

    return jsonify({
        'registros': registros,
        'totales':   {'cantidad': len(registros), 'monto': round(monto_total, 2)},
        'por_mes':   por_mes,
    })


@app.route('/api/reportes/exportar')
@login_required
def api_reportes_exportar():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    desde      = request.args.get('desde', '').strip()
    hasta      = request.args.get('hasta', '').strip()
    cliente    = request.args.get('cliente', '').strip()

    empresa = EmpresaRepository.get_by_id(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403

    registros = _leer_registros_reporte(empresa_id, desde, hasta, cliente)
    if not registros:
        return 'No hay datos para exportar', 404

    df = pd.DataFrame(registros, columns=[
        'mes', 'fecha', 'razon_social', 'punto_venta',
        'nro_cbte', 'tipo_cbte', 'imp_neto', 'imp_iva', 'imp_total', 'cae'
    ])
    df.columns = ['Mes', 'Fecha', 'Cliente', 'Pto. Venta',
                  'Nro. Cbte', 'Tipo Cbte', 'Neto', 'IVA', 'Total', 'CAE']

    # Fila de totales
    total_row = pd.DataFrame([{
        'Mes': '', 'Fecha': '', 'Cliente': 'TOTAL',
        'Pto. Venta': '', 'Nro. Cbte': len(registros), 'Tipo Cbte': '',
        'Neto': df['Neto'].sum(), 'IVA': df['IVA'].sum(),
        'Total': df['Total'].sum(), 'CAE': '',
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
