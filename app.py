import functools
import io
import json
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
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import config
import wsaa
import wsfe
import factura_pdf
from openssl_util import encontrar_openssl

app = Flask(__name__)
app.secret_key = 'arca_2026_secret_key_change_in_prod'

BASE     = os.path.dirname(os.path.abspath(__file__))
UPLOAD   = os.path.join(BASE, 'uploads')
CERTS    = os.path.join(BASE, 'certificados')
EMPRESAS = os.path.join(BASE, 'empresas.json')
USUARIOS = os.path.join(BASE, 'usuarios.json')
COLUMNAS = [
    'punto_venta', 'tipo_cbte', 'concepto',
    'doc_tipo', 'doc_nro', 'razon_social',
    'fecha', 'imp_neto', 'alicuota', 'imp_iva', 'imp_total',
]

os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(CERTS,  exist_ok=True)


# ---------- helpers empresas --------------------------------------------------

def _load_empresas():
    if not os.path.exists(EMPRESAS):
        return []
    with open(EMPRESAS, 'r', encoding='utf-8') as f:
        return json.load(f)

def _save_empresas(data):
    with open(EMPRESAS, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _get_empresa(empresa_id):
    for e in _load_empresas():
        if e['id'] == empresa_id:
            return e
    return None

def _upload_path(empresa_id):
    return os.path.join(UPLOAD, f'facturas_{empresa_id}.xlsx')

def _resultado_path(empresa_id):
    return os.path.join(UPLOAD, f'facturas_{empresa_id}_resultado.xlsx')

def _empresa_urls(empresa):
    if empresa.get('homologacion'):
        return config.WSAA_URL_HOMO, config.WSFE_WSDL_HOMO
    return config.WSAA_URL_PROD, config.WSFE_WSDL_PROD


# ---------- helpers usuarios --------------------------------------------------

def _load_usuarios():
    if not os.path.exists(USUARIOS):
        return []
    with open(USUARIOS, 'r', encoding='utf-8') as f:
        return json.load(f)

def _save_usuarios(data):
    with open(USUARIOS, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _get_current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    for u in _load_usuarios():
        if u['id'] == uid:
            return u
    return None

def _user_empresas(user):
    """Retorna las empresas accesibles para el usuario."""
    todas = _load_empresas()
    if user['rol'] == 'admin' or not user.get('empresas'):
        return todas
    return [e for e in todas if e['id'] in user.get('empresas', [])]

def _user_can_access(user, empresa_id):
    """Verifica si el usuario puede operar sobre esa empresa."""
    if user['rol'] == 'admin':
        return True
    return empresa_id in user.get('empresas', [])

def _init_admin():
    """Crea el usuario admin por defecto si no hay ningún usuario registrado."""
    if not _load_usuarios():
        _save_usuarios([{
            'id':            'admin',
            'username':      'admin',
            'password_hash': generate_password_hash('admin123'),
            'rol':           'admin',
            'nombre':        'Administrador',
            'empresas':      [],
        }])
        print("Usuario admin creado con contraseña: admin123 — ¡cambiarla desde Admin!")

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
        usuario  = next((u for u in _load_usuarios() if u['username'] == username), None)
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
    user = _get_current_user()
    return jsonify(_user_empresas(user))

@app.route('/api/empresas', methods=['POST'])
@admin_required
def api_empresa_add():
    data   = request.get_json(force=True)
    nombre = (data.get('nombre') or '').strip()
    cuit   = (data.get('cuit')   or '').strip()
    cert   = (data.get('cert')   or '').strip()
    key    = (data.get('key')    or '').strip()
    homo   = bool(data.get('homologacion', False))

    if not nombre or not cuit or not cert or not key:
        return jsonify({'error': 'Nombre, CUIT, certificado y clave son obligatorios'}), 400
    if not re.fullmatch(r'\d{11}', cuit):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400

    empresas = _load_empresas()
    if any(e['cuit'] == cuit for e in empresas):
        return jsonify({'error': f'Ya existe una empresa con CUIT {cuit}'}), 400

    empresa_id = re.sub(r'[^a-z0-9]', '', nombre.lower())[:20] or str(uuid.uuid4())[:8]
    if any(e['id'] == empresa_id for e in empresas):
        empresa_id = empresa_id + '_' + str(uuid.uuid4())[:4]

    empresas.append({
        'id':                 empresa_id,
        'nombre':             nombre,
        'cuit':               cuit,
        'cert':               cert,
        'key':                key,
        'homologacion':       homo,
        'domicilio':          (data.get('domicilio') or '').strip(),
        'telefono':           (data.get('telefono')  or '').strip(),
        'localidad':          (data.get('localidad') or '').strip(),
        'ing_brutos':         (data.get('ing_brutos') or '').strip(),
        'inicio_actividades': (data.get('inicio_actividades') or '').strip(),
    })
    _save_empresas(empresas)
    return jsonify({'ok': True, 'id': empresa_id})

@app.route('/api/empresas/<empresa_id>', methods=['PUT'])
@admin_required
def api_empresa_edit(empresa_id):
    empresas = _load_empresas()
    idx = next((i for i, e in enumerate(empresas) if e['id'] == empresa_id), None)
    if idx is None:
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

    if any(e['cuit'] == cuit and e['id'] != empresa_id for e in empresas):
        return jsonify({'error': f'Ya existe otra empresa con CUIT {cuit}'}), 400

    empresas[idx].update({
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
    _save_empresas(empresas)
    return jsonify({'ok': True})

@app.route('/api/empresas/<empresa_id>', methods=['DELETE'])
@admin_required
def api_empresa_delete(empresa_id):
    empresas = _load_empresas()
    nuevas   = [e for e in empresas if e['id'] != empresa_id]
    if len(nuevas) == len(empresas):
        return jsonify({'error': 'Empresa no encontrada'}), 404
    _save_empresas(nuevas)
    for path in [_upload_path(empresa_id), _resultado_path(empresa_id)]:
        if os.path.exists(path):
            os.remove(path)
    return jsonify({'ok': True})


# ---------- API usuarios ------------------------------------------------------

@app.route('/api/usuarios')
@admin_required
def api_usuarios():
    usuarios = _load_usuarios()
    # No devolver password_hash al frontend
    return jsonify([{k: v for k, v in u.items() if k != 'password_hash'}
                    for u in usuarios])

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

    usuarios = _load_usuarios()
    if any(u['username'] == username for u in usuarios):
        return jsonify({'error': f'Ya existe un usuario con ese nombre'}), 400

    uid = re.sub(r'[^a-z0-9]', '', username.lower())[:20] or str(uuid.uuid4())[:8]
    if any(u['id'] == uid for u in usuarios):
        uid = uid + '_' + str(uuid.uuid4())[:4]

    usuarios.append({
        'id':            uid,
        'username':      username,
        'password_hash': generate_password_hash(password),
        'rol':           rol,
        'nombre':        nombre or username,
        'empresas':      empresas if rol == 'usuario' else [],
    })
    _save_usuarios(usuarios)
    return jsonify({'ok': True, 'id': uid})

@app.route('/api/usuarios/<uid>', methods=['PUT'])
@admin_required
def api_usuario_edit(uid):
    usuarios = _load_usuarios()
    idx = next((i for i, u in enumerate(usuarios) if u['id'] == uid), None)
    if idx is None:
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
    if any(u['username'] == username and u['id'] != uid for u in usuarios):
        return jsonify({'error': 'Ese nombre de usuario ya está en uso'}), 400

    usuarios[idx].update({
        'username': username,
        'nombre':   nombre or username,
        'rol':      rol,
        'empresas': empresas if rol == 'usuario' else [],
    })
    if password:  # solo actualizar si se envió nueva contraseña
        usuarios[idx]['password_hash'] = generate_password_hash(password)

    _save_usuarios(usuarios)
    return jsonify({'ok': True})

@app.route('/api/usuarios/<uid>', methods=['DELETE'])
@admin_required
def api_usuario_delete(uid):
    user = _get_current_user()
    if user['id'] == uid:
        return jsonify({'error': 'No podés eliminar tu propio usuario'}), 400
    usuarios  = _load_usuarios()
    nuevos    = [u for u in usuarios if u['id'] != uid]
    if len(nuevos) == len(usuarios):
        return jsonify({'error': 'Usuario no encontrado'}), 404
    _save_usuarios(nuevos)
    return jsonify({'ok': True})

@app.route('/api/usuarios/<uid>/cambiar-password', methods=['POST'])
@login_required
def api_cambiar_password(uid):
    """Permite a un usuario cambiar su propia contraseña."""
    user = _get_current_user()
    # Solo el propio usuario o un admin pueden cambiar la contraseña
    if user['id'] != uid and user['rol'] != 'admin':
        return jsonify({'error': 'Sin permisos'}), 403

    data         = request.get_json(force=True)
    nueva        = (data.get('nueva') or '').strip()
    confirmacion = (data.get('confirmacion') or '').strip()

    if not nueva or len(nueva) < 6:
        return jsonify({'error': 'La contraseña debe tener al menos 6 caracteres'}), 400
    if nueva != confirmacion:
        return jsonify({'error': 'Las contraseñas no coinciden'}), 400

    usuarios = _load_usuarios()
    idx = next((i for i, u in enumerate(usuarios) if u['id'] == uid), None)
    if idx is None:
        return jsonify({'error': 'Usuario no encontrado'}), 404

    usuarios[idx]['password_hash'] = generate_password_hash(nueva)
    _save_usuarios(usuarios)
    return jsonify({'ok': True})


# ---------- flujo de facturación ----------------------------------------------

@app.route('/upload', methods=['POST'])
@login_required
def upload():
    user       = _get_current_user()
    empresa_id = request.form.get('empresa_id', '').strip()
    empresa    = _get_empresa(empresa_id)
    if not empresa:
        return jsonify({'error': 'Seleccioná una empresa antes de cargar el archivo'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'No tenés acceso a esta empresa'}), 403

    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No se seleccionó archivo'}), 400
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    path = _upload_path(empresa_id)
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
    empresa    = _get_empresa(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400
    if not _user_can_access(user, empresa_id):
        return jsonify({'error': 'No tenés acceso a esta empresa'}), 403

    path = _upload_path(empresa_id)
    if not os.path.exists(path):
        return jsonify({'error': 'No hay archivo cargado para esta empresa'}), 400

    wsaa_url, wsfe_wsdl = _empresa_urls(empresa)

    cert_path = empresa.get('cert', '')
    key_path  = empresa.get('key', '')
    if not os.path.isfile(cert_path):
        return jsonify({'error': f'Certificado no encontrado: {cert_path}\nAndá a Admin y actualizá la ruta del certificado (.crt)'}), 400
    if not os.path.isfile(key_path):
        return jsonify({'error': f'Clave privada no encontrada: {key_path}\nAndá a Admin y actualizá la ruta de la clave (.key)'}), 400

    try:
        token, sign = wsaa.get_ticket(
            'wsfe', cert_path, key_path, wsaa_url, empresa['cuit']
        )
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
                        'fila': idx + 2, 'nro': nro,
                        'resultado': 'APROBADO',
                        'cae': det.CAE, 'vto_cae': str(det.CAEFchVto), 'obs': '',
                    })
                else:
                    obs = ''
                    if det.Observaciones:
                        obs = '; '.join(o.Msg for o in det.Observaciones.Obs)
                    resultados.append({
                        'fila': idx + 2, 'nro': nro,
                        'resultado': 'RECHAZADO',
                        'cae': '', 'vto_cae': '', 'obs': obs,
                    })

            except Exception as e:
                resultados.append({
                    'fila': idx + 2, 'nro': 0,
                    'resultado': 'ERROR',
                    'cae': '', 'vto_cae': '', 'obs': str(e),
                })

        _guardar_resultado(path, _resultado_path(empresa_id), resultados)

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
        tb = traceback.format_exc()
        print(f"\n=== ERROR /procesar ===\n{tb}\n======================\n")
        return jsonify({'error': str(e)}), 500


def _guardar_resultado(src_path, dest_path, resultados):
    wb = load_workbook(src_path)
    ws = wb.active
    lc = ws.max_column + 1

    for i, h in enumerate(['Nro_Cbte', 'Resultado', 'CAE', 'Vto_CAE', 'Observaciones']):
        ws.cell(1, lc + i, h)

    verde    = PatternFill(fill_type='solid', fgColor='C6EFCE')
    rojo     = PatternFill(fill_type='solid', fgColor='FFC7CE')
    amarillo = PatternFill(fill_type='solid', fgColor='FFEB9C')

    for r in resultados:
        rn = r['fila']
        ws.cell(rn, lc,     r['nro'])
        ws.cell(rn, lc + 1, r['resultado'])
        ws.cell(rn, lc + 2, r['cae'])
        ws.cell(rn, lc + 3, r['vto_cae'])
        ws.cell(rn, lc + 4, r['obs'])
        fill = verde if r['resultado'] == 'APROBADO' else (rojo if r['resultado'] == 'RECHAZADO' else amarillo)
        for col in range(1, lc + 5):
            ws.cell(rn, col).fill = fill

    wb.save(dest_path)


@app.route('/descargar')
@login_required
def descargar():
    user       = _get_current_user()
    empresa_id = request.args.get('empresa_id', '').strip()
    empresa    = _get_empresa(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403
    path = _resultado_path(empresa_id)
    if not os.path.exists(path):
        return 'No hay resultado disponible', 404
    nombre_archivo = f'facturas_{empresa["cuit"]}_resultado.xlsx'
    return send_file(path, as_attachment=True, download_name=nombre_archivo)


@app.route('/plantilla')
@login_required
def plantilla():
    df = pd.DataFrame([{
        'punto_venta': 6,
        'tipo_cbte':   11,
        'concepto':    2,
        'doc_tipo':    99,
        'doc_nro':     0,
        'razon_social': 'Consumidor Final',
        'fecha':       datetime.today().strftime('%Y-%m-%d'),
        'imp_neto':    1000.00,
        'alicuota':    0,
        'imp_iva':     0.00,
        'imp_total':   1000.00,
    }])
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Facturas')
    out.seek(0)
    return send_file(out, as_attachment=True, download_name='plantilla_facturas.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- generador de PDF -------------------------------------------------

@app.route('/pdf/<empresa_id>/<int:fila>')
@login_required
def pdf_desde_resultado(empresa_id, fila):
    user    = _get_current_user()
    empresa = _get_empresa(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    if not _user_can_access(user, empresa_id):
        return 'Acceso denegado', 403

    path = _resultado_path(empresa_id)
    if not os.path.exists(path):
        return 'No hay resultados guardados para esta empresa', 404

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]

        idx = fila - 2
        if idx < 0 or idx >= len(df):
            return f'Fila {fila} no encontrada', 404

        row      = df.iloc[idx].to_dict()
        registro = {
            'punto_venta': row.get('punto_venta', 0),
            'tipo_cbte':   row.get('tipo_cbte', 11),
            'concepto':    row.get('concepto', 2),
            'doc_tipo':    row.get('doc_tipo', 99),
            'doc_nro':     row.get('doc_nro', 0),
            'razon_social':row.get('razon_social', ''),
            'fecha':       str(row.get('fecha', '')),
            'imp_neto':    row.get('imp_neto', 0),
            'alicuota':    row.get('alicuota', 0),
            'imp_iva':     row.get('imp_iva', 0),
            'imp_total':   row.get('imp_total', 0),
        }
        resultado = {
            'nro':     int(row.get('nro_cbte', 0)),
            'cae':     str(row.get('cae', '')),
            'vto_cae': str(row.get('vto_cae', '')),
        }

        pdf_buf = factura_pdf.generar_pdf(empresa, registro, resultado)
        pv      = int(registro['punto_venta'])
        nro     = int(resultado['nro'])
        nombre  = f"factura_{pv:04d}-{nro:08d}.pdf"
        return send_file(pdf_buf, as_attachment=False,
                         download_name=nombre, mimetype='application/pdf')
    except Exception as e:
        return f'Error al generar PDF: {e}', 500


@app.route('/imprimir', methods=['POST'])
@login_required
def imprimir():
    user       = _get_current_user()
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    empresa    = _get_empresa(empresa_id)
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
        nro     = int(resultado.get('nro', 0))
        pv      = int(registro.get('punto_venta', 0))
        nombre  = f"factura_{pv:04d}-{nro:08d}.pdf"
        return send_file(pdf_buf, as_attachment=False,
                         download_name=nombre, mimetype='application/pdf')
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
        r1 = subprocess.run(
            [openssl, 'genrsa', '-out', key_path, '2048'],
            capture_output=True
        )
        if r1.returncode != 0:
            raise Exception(r1.stderr.decode('utf-8', errors='replace'))

        subject = f'/C=AR/O={nombre}/serialNumber=CUIT {cuit}/CN={cuit}'
        if email:
            subject += f'/emailAddress={email}'

        r2 = subprocess.run(
            [openssl, 'req', '-new', '-key', key_path, '-out', csr_path, '-subj', subject],
            capture_output=True
        )
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


# ---------- página de administración -----------------------------------------

@app.route('/admin')
@admin_required
def admin():
    user     = _get_current_user()
    empresas = _load_empresas()
    usuarios = [{k: v for k, v in u.items() if k != 'password_hash'}
                for u in _load_usuarios()]
    return render_template('admin.html', empresas=empresas, usuarios=usuarios,
                           certs_dir=CERTS, current_user=user)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
