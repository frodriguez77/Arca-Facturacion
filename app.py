import io
import os
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import config
import wsaa
import wsfe

app = Flask(__name__)
app.secret_key = 'arca_2026'

UPLOAD  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
COLUMNAS = [
    'punto_venta', 'tipo_cbte', 'concepto',
    'doc_tipo', 'doc_nro', 'razon_social',
    'fecha', 'imp_neto', 'alicuota', 'imp_iva', 'imp_total',
]


@app.route('/')
def index():
    return render_template('index.html', cuit=config.CUIT, homo=config.HOMOLOGACION)


@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No se seleccionó archivo'}), 400
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx'}), 400

    path = os.path.join(UPLOAD, 'facturas.xlsx')
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
    path = os.path.join(UPLOAD, 'facturas.xlsx')
    if not os.path.exists(path):
        return jsonify({'error': 'No hay archivo cargado'}), 400

    try:
        token, sign = wsaa.get_ticket('wsfe', config.CERT, config.KEY, config.WSAA_URL)
        auth        = {'Token': token, 'Sign': sign, 'Cuit': int(config.CUIT)}
        client      = wsfe.get_client(config.WSFE_WSDL)

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

                result = wsfe.procesar_comprobante(client, auth, config.CUIT, pv, tipo, comp, nro)
                det    = result.FeDetResp.FECAEDetResponse[0]

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

        _guardar_resultado(path, resultados)

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
        return jsonify({'error': str(e)}), 500


def _guardar_resultado(src_path, resultados):
    wb = load_workbook(src_path)
    ws = wb.active
    lc = ws.max_column + 1

    for i, h in enumerate(['Nro_Cbte', 'Resultado', 'CAE', 'Vto_CAE', 'Observaciones']):
        ws.cell(1, lc + i, h)

    verde   = PatternFill(fill_type='solid', fgColor='C6EFCE')
    rojo    = PatternFill(fill_type='solid', fgColor='FFC7CE')
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

    wb.save(os.path.join(UPLOAD, 'facturas_resultado.xlsx'))


@app.route('/descargar')
def descargar():
    path = os.path.join(UPLOAD, 'facturas_resultado.xlsx')
    if not os.path.exists(path):
        return 'No hay resultado disponible', 404
    return send_file(path, as_attachment=True, download_name='facturas_resultado.xlsx')


@app.route('/plantilla')
def plantilla():
    df = pd.DataFrame([{
        'punto_venta': 6,
        'tipo_cbte':   11,       # Factura C - Monotributista
        'concepto':    2,        # Servicios
        'doc_tipo':    99,       # Consumidor Final
        'doc_nro':     0,
        'razon_social': 'Consumidor Final',
        'fecha':       datetime.today().strftime('%Y-%m-%d'),
        'imp_neto':    1000.00,  # En FC imp_neto = imp_total
        'alicuota':    0,        # Sin IVA
        'imp_iva':     0.00,
        'imp_total':   1000.00,
    }])
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Facturas')
    out.seek(0)
    return send_file(out, as_attachment=True, download_name='plantilla_facturas.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    app.run(debug=True, port=5000)
