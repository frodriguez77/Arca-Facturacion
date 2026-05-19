import io
import json
import os
import re
import subprocess
import traceback
import uuid
import zipfile
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import config
import wsaa
import wsfe
import factura_pdf
from openssl_util import encontrar_openssl

app = Flask(__name__)
app.secret_key = 'arca_2026'

BASE     = os.path.dirname(os.path.abspath(__file__))
UPLOAD   = os.path.join(BASE, 'uploads')
CERTS    = os.path.join(BASE, 'certificados')
EMPRESAS = os.path.join(BASE, 'empresas.json')
COLUMNAS = [
    'punto_venta', 'tipo_cbte', 'concepto',
    'doc_tipo', 'doc_nro', 'razon_social',
    'fecha', 'imp_neto', 'alicuota', 'imp_iva', 'imp_total',
]

os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(CERTS,  exist_ok=True)


# ---------- helpers ----------------------------------------------------------

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


# ---------- vistas principales -----------------------------------------------

@app.route('/')
def index():
    empresas = _load_empresas()
    return render_template('index.html', empresas=empresas)


# ---------- API empresas ------------------------------------------------------

@app.route('/api/empresas')
def api_empresas():
    return jsonify(_load_empresas())

@app.route('/api/empresas', methods=['POST'])
def api_empresa_add():
    data = request.get_json(force=True)
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
def api_empresa_edit(empresa_id):
    empresas = _load_empresas()
    idx = next((i for i,e in enumerate(empresas) if e['id'] == empresa_id), None)
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

    # Si cambió el CUIT verificar que no exista en otra empresa
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
def api_empresa_delete(empresa_id):
    empresas = _load_empresas()
    nuevas = [e for e in empresas if e['id'] != empresa_id]
    if len(nuevas) == len(empresas):
        return jsonify({'error': 'Empresa no encontrada'}), 404
    _save_empresas(nuevas)
    for path in [_upload_path(empresa_id), _resultado_path(empresa_id)]:
        if os.path.exists(path):
            os.remove(path)
    return jsonify({'ok': True})


# ---------- flujo de facturación ----------------------------------------------

@app.route('/upload', methods=['POST'])
def upload():
    empresa_id = request.form.get('empresa_id', '').strip()
    empresa = _get_empresa(empresa_id)
    if not empresa:
        return jsonify({'error': 'Seleccioná una empresa antes de cargar el archivo'}), 400

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
def procesar():
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    empresa    = _get_empresa(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

    path = _upload_path(empresa_id)
    if not os.path.exists(path):
        return jsonify({'error': 'No hay archivo cargado para esta empresa'}), 400

    wsaa_url, wsfe_wsdl = _empresa_urls(empresa)

    # Verificar que existan los archivos de certificado antes de llamar a AFIP
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
def descargar():
    empresa_id = request.args.get('empresa_id', '').strip()
    empresa    = _get_empresa(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404
    path = _resultado_path(empresa_id)
    if not os.path.exists(path):
        return 'No hay resultado disponible', 404
    nombre_archivo = f'facturas_{empresa["cuit"]}_resultado.xlsx'
    return send_file(path, as_attachment=True, download_name=nombre_archivo)


@app.route('/plantilla')
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
def pdf_desde_resultado(empresa_id, fila):
    """Genera el PDF leyendo el Excel de resultados ya guardado en disco."""
    empresa = _get_empresa(empresa_id)
    if not empresa:
        return 'Empresa no encontrada', 404

    path = _resultado_path(empresa_id)
    if not os.path.exists(path):
        return 'No hay resultados guardados para esta empresa', 404

    try:
        df = pd.read_excel(path)
        df.columns = [c.lower().strip().replace(' ', '_') for c in df.columns]

        # fila 2 en Excel = índice 0 en DataFrame
        idx = fila - 2
        if idx < 0 or idx >= len(df):
            return f'Fila {fila} no encontrada', 404

        row = df.iloc[idx].to_dict()

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
        pv  = int(registro['punto_venta'])
        nro = int(resultado['nro'])
        nombre = f"factura_{pv:04d}-{nro:08d}.pdf"
        return send_file(pdf_buf, as_attachment=False,
                         download_name=nombre, mimetype='application/pdf')
    except Exception as e:
        return f'Error al generar PDF: {e}', 500


@app.route('/imprimir', methods=['POST'])
def imprimir():
    data       = request.get_json(force=True)
    empresa_id = (data.get('empresa_id') or '').strip()
    empresa    = _get_empresa(empresa_id)
    if not empresa:
        return jsonify({'error': 'Empresa no encontrada'}), 400

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
def generar_csr():
    data   = request.get_json(force=True)
    cuit   = (data.get('cuit')   or '').strip()
    nombre = (data.get('nombre') or '').strip()
    email  = (data.get('email')  or '').strip()

    if not cuit or not nombre:
        return jsonify({'error': 'CUIT y nombre son obligatorios'}), 400
    if not re.fullmatch(r'\d{11}', cuit):
        return jsonify({'error': 'El CUIT debe tener 11 dígitos sin guiones'}), 400

    # Subcarpeta por CUIT para mantener ordenado
    cuit_dir = os.path.join(CERTS, cuit)
    os.makedirs(cuit_dir, exist_ok=True)
    key_path = os.path.join(cuit_dir, f'{cuit}_clave.key')
    csr_path = os.path.join(cuit_dir, f'{cuit}.csr')

    try:
        openssl = encontrar_openssl()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    try:
        # 1. Generar clave privada RSA 2048
        r1 = subprocess.run(
            [openssl, 'genrsa', '-out', key_path, '2048'],
            capture_output=True
        )
        if r1.returncode != 0:
            raise Exception(r1.stderr.decode('utf-8', errors='replace'))

        # 2. Generar CSR con los datos requeridos por AFIP
        subject = f'/C=AR/O={nombre}/serialNumber=CUIT {cuit}/CN={cuit}'
        if email:
            subject += f'/emailAddress={email}'

        r2 = subprocess.run(
            [openssl, 'req', '-new',
             '-key',     key_path,
             '-out',     csr_path,
             '-subj',    subject],
            capture_output=True
        )
        if r2.returncode != 0:
            raise Exception(r2.stderr.decode('utf-8', errors='replace'))

        # 3. Empaquetar .key y .csr en un ZIP para descargar
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(key_path, f'{cuit}_clave.key')
            zf.write(csr_path, f'{cuit}.csr')
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=f'certificado_afip_{cuit}.zip',
            mimetype='application/zip',
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ---------- página de administración -----------------------------------------

@app.route('/admin')
def admin():
    empresas = _load_empresas()
    return render_template('admin.html', empresas=empresas, certs_dir=CERTS)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
