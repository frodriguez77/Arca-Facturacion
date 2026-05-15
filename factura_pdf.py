"""
Generador de PDF de factura electrónica con QR AFIP (RG 4291).
"""
import base64
import io
import json
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
import qrcode

W, H = A4
ML = MR = 2 * cm
MT = MB = 2 * cm
CW = W - ML - MR   # ancho útil ≈ 17 cm

# ---------- tablas de referencia ---------------------------------------------

TIPO_NOMBRE = {
    1: 'Factura A',         2: 'Nota de Débito A',  3: 'Nota de Crédito A',
    6: 'Factura B',         7: 'Nota de Débito B',  8: 'Nota de Crédito B',
    11: 'Factura C',        12: 'Nota de Débito C', 13: 'Nota de Crédito C',
}
TIPO_LETRA = {
    1: 'A', 2: 'A', 3: 'A',
    6: 'B', 7: 'B', 8: 'B',
    11: 'C', 12: 'C', 13: 'C',
}
DOC_NOMBRE = {
    80: 'CUIT', 86: 'CUIL', 87: 'CDI',
    96: 'DNI',  99: 'Consumidor Final',
}
CONCEPTO_NOMBRE = {1: 'Productos', 2: 'Servicios', 3: 'Productos y Servicios'}
IVA_COND = {
    'A': 'Responsable Inscripto',
    'B': 'Responsable Inscripto',
    'C': 'Monotributista / Consumidor Final',
}


# ---------- helpers ----------------------------------------------------------

def _fmt(v):
    """$ 1.000,50 — formato argentino."""
    v = round(float(v or 0), 2)
    neg = v < 0
    av = abs(v)
    i = int(av)
    d = round((av - i) * 100)
    s = f"$ {i:,}".replace(",", ".") + f",{d:02d}"
    return f"-{s}" if neg else s


def _parse_fecha(raw):
    """Convierte fecha en varios formatos a datetime."""
    raw = str(raw).strip()
    for fmt in ('%Y-%m-%d', '%Y%m%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(raw[:len(fmt.replace('%Y','0000').replace('%m','00').replace('%d','00').replace('%H','00').replace('%M','00').replace('%S','00'))], fmt)
        except Exception:
            pass
    # fallback
    for fmt in ('%Y-%m-%d', '%Y%m%d'):
        try:
            return datetime.strptime(raw[:10] if '-' in raw else raw[:8], fmt)
        except Exception:
            pass
    return None


def _fmt_fecha(raw):
    dt = _parse_fecha(raw)
    return dt.strftime('%d/%m/%Y') if dt else str(raw)


def _cuit_fmt(cuit):
    c = str(cuit)
    if len(c) == 11:
        return f"{c[:2]}-{c[2:10]}-{c[10]}"
    return c


def _qr_image(cuit, pv, tipo, nro, fecha_raw, total, doc_tipo, doc_nro, cae):
    dt = _parse_fecha(fecha_raw)
    fecha_iso = dt.strftime('%Y-%m-%d') if dt else str(fecha_raw)[:10]

    data = {
        "ver":        1,
        "fecha":      fecha_iso,
        "cuit":       int(cuit),
        "ptoVta":     int(pv),
        "tipoCmp":    int(tipo),
        "nroCmp":     int(nro),
        "importe":    round(float(total), 2),
        "moneda":     "PES",
        "ctz":        1,
        "tipoDocRec": int(doc_tipo),
        "nroDocRec":  int(str(doc_nro).split('.')[0] or 0),
        "tipoCodAut": "E",
        "codAut":     int(cae),
    }
    b64 = base64.b64encode(
        json.dumps(data, separators=(',', ':')).encode()
    ).decode()
    url = f"https://www.afip.gob.ar/fe/qr/?p={b64}"

    qr = qrcode.QRCode(version=1, box_size=5, border=2,
                       error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


# ---------- generador principal ----------------------------------------------

def generar_pdf(empresa, registro, resultado):
    """
    Retorna un BytesIO con el PDF listo para enviar al navegador.

    empresa  : dict  {nombre, cuit, homologacion}
    registro : dict  campos del Excel (punto_venta, tipo_cbte, concepto,
                     doc_tipo, doc_nro, razon_social, fecha,
                     imp_neto, alicuota, imp_iva, imp_total)
    resultado: dict  {nro, cae, vto_cae}
    """
    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)

    # -- extraer y convertir datos --
    cuit      = str(empresa['cuit'])
    nom_emp   = empresa['nombre']
    homo      = empresa.get('homologacion', False)

    tipo_cbte = int(registro['tipo_cbte'])
    pto_venta = int(registro['punto_venta'])
    concepto  = int(registro['concepto'])
    doc_tipo  = int(registro['doc_tipo'])
    doc_nro   = str(registro['doc_nro']).split('.')[0]
    razon_soc = str(registro['razon_social'])
    fecha_raw = str(registro['fecha'])
    imp_neto  = float(registro.get('imp_neto', 0))
    alicuota  = float(registro.get('alicuota', 0))
    imp_iva   = float(registro.get('imp_iva', 0))
    imp_total = float(registro.get('imp_total', 0))

    nro_cbte  = int(resultado['nro'])
    cae       = str(resultado['cae'])
    vto_cae   = str(resultado.get('vto_cae', ''))

    letra    = TIPO_LETRA.get(tipo_cbte, '?')
    tipo_nom = TIPO_NOMBRE.get(tipo_cbte, f'Tipo {tipo_cbte}')
    nro_fmt  = f"{pto_venta:04d}-{nro_cbte:08d}"
    doc_nom  = DOC_NOMBRE.get(doc_tipo, f'Doc.{doc_tipo}')
    conc_nom = CONCEPTO_NOMBRE.get(concepto, str(concepto))
    iva_cond = IVA_COND.get(letra, '')

    fecha_disp = _fmt_fecha(fecha_raw)
    vto_disp   = _fmt_fecha(vto_cae)

    AZUL  = colors.HexColor('#1d4ed8')
    VERDE = colors.HexColor('#065f46')
    GRIS  = colors.HexColor('#6b7280')
    BGRIS = colors.HexColor('#f3f4f6')

    # -- helpers de canvas --
    def hline(y, x1=ML, x2=ML+CW, w=0.5, color=colors.black):
        c.setStrokeColor(color)
        c.setLineWidth(w)
        c.line(x1, y, x2, y)

    def box(x, y, bw, bh, fill_color=None, stroke_color=colors.black, lw=0.5):
        c.setLineWidth(lw)
        c.setStrokeColor(stroke_color)
        if fill_color:
            c.setFillColor(fill_color)
            c.rect(x, y, bw, bh, fill=1, stroke=1)
            c.setFillColor(colors.black)
        else:
            c.rect(x, y, bw, bh, fill=0, stroke=1)

    def txt(x, y, s, size=9, bold=False, color=colors.black, align='left'):
        c.setFillColor(color)
        c.setFont('Helvetica-Bold' if bold else 'Helvetica', size)
        s = str(s)
        if align == 'right':
            c.drawRightString(x, y, s)
        elif align == 'center':
            c.drawCentredString(x, y, s)
        else:
            c.drawString(x, y, s)
        c.setFillColor(colors.black)

    # =========================================================================
    # ENCABEZADO
    # =========================================================================
    y_top    = H - MT
    hdr_h    = 5.2 * cm
    y_hdr_b  = y_top - hdr_h

    box(ML, y_hdr_b, CW, hdr_h)

    # divisores verticales del cuadro de letra
    xll = ML + CW / 2 - 1.4 * cm
    xlr = ML + CW / 2 + 1.4 * cm
    c.setLineWidth(0.5)
    c.line(xll, y_hdr_b, xll, y_top)
    c.line(xlr, y_hdr_b, xlr, y_top)

    # letra grande en el centro
    cx = ML + CW / 2
    txt(cx, y_hdr_b + hdr_h - 2.0 * cm, letra, size=64, bold=True, align='center')
    txt(cx, y_hdr_b + 0.85 * cm, 'ORIGINAL', size=8, align='center')
    if homo:
        txt(cx, y_hdr_b + 0.35 * cm, 'HOMOLOGACIÓN', size=7, color=colors.red, align='center')

    # empresa — columna izquierda
    xl = ML + 0.4 * cm
    yt = y_top - 0.6 * cm
    txt(xl, yt,             nom_emp, size=11, bold=True)
    txt(xl, yt - 0.55*cm,  f'CUIT: {_cuit_fmt(cuit)}', size=8)
    txt(xl, yt - 1.05*cm,  f'Condición IVA: {iva_cond}', size=8)

    # comprobante — columna derecha
    xr_end = ML + CW - 0.4 * cm
    txt(xr_end, yt,             tipo_nom, size=10, bold=True, align='right')
    txt(xr_end, yt - 0.55*cm,  f'Nro: {nro_fmt}', size=9, align='right')
    txt(xr_end, yt - 1.05*cm,  f'Fecha: {fecha_disp}', size=9, align='right')

    # =========================================================================
    # DATOS DEL RECEPTOR
    # =========================================================================
    y_rec_t = y_hdr_b - 0.25 * cm
    rec_h   = 2.0 * cm
    y_rec_b = y_rec_t - rec_h

    box(ML, y_rec_b, CW, rec_h, fill_color=BGRIS)

    xl2  = ML + 0.4 * cm
    col2 = ML + CW / 2
    yr1  = y_rec_t - 0.55 * cm
    yr2  = y_rec_t - 1.10 * cm

    txt(xl2,  y_rec_t - 0.2*cm, 'RECEPTOR', size=7, bold=True, color=GRIS)
    txt(xl2,  yr1,  f'Razón Social: {razon_soc}', size=9)
    txt(col2, yr1,  f'{doc_nom}: {doc_nro}', size=9)
    txt(xl2,  yr2,  f'Concepto: {conc_nom}', size=9)

    # =========================================================================
    # IMPORTES
    # =========================================================================
    y_imp_t = y_rec_b - 0.25 * cm
    imp_h   = 3.2 * cm
    y_imp_b = y_imp_t - imp_h

    box(ML, y_imp_b, CW, imp_h)

    # cabecera tabla
    box(ML, y_imp_t - 0.65*cm, CW, 0.65*cm, fill_color=BGRIS)
    txt(ML + 0.4*cm,      y_imp_t - 0.45*cm, 'DESCRIPCIÓN', size=8, bold=True)
    txt(ML+CW - 0.4*cm,   y_imp_t - 0.45*cm, 'IMPORTE', size=8, bold=True, align='right')

    yi    = y_imp_t - 1.2 * cm
    step  = 0.55 * cm
    xval  = ML + CW - 0.4 * cm

    # fila neto
    neto_lbl = 'Importe Neto' if tipo_cbte not in [11, 12, 13] else 'Importe Total'
    txt(ML + 0.4*cm, yi, neto_lbl, size=9)
    txt(xval, yi, _fmt(imp_neto), size=9, align='right')
    yi -= step

    # fila IVA
    if imp_iva > 0:
        iva_lbl = f'IVA {alicuota:.0f}%' if alicuota > 0 else 'IVA'
        txt(ML + 0.4*cm, yi, iva_lbl, size=9)
        txt(xval, yi, _fmt(imp_iva), size=9, align='right')
        yi -= step

    # línea y total
    hline(yi + 0.4*cm, x1=ML + CW/2, x2=ML+CW - 0.2*cm)
    txt(ML + 0.4*cm, yi, 'TOTAL', size=11, bold=True, color=AZUL)
    txt(xval, yi, _fmt(imp_total), size=11, bold=True, color=AZUL, align='right')

    # =========================================================================
    # CAE + QR
    # =========================================================================
    y_cae_t = y_imp_b - 0.25 * cm
    cae_h   = 5.2 * cm
    y_cae_b = y_cae_t - cae_h

    box(ML, y_cae_b, CW, cae_h)

    # QR (derecha)
    qr_sz = 4.0 * cm
    qr_x  = ML + CW - qr_sz - 0.5 * cm
    qr_y  = y_cae_b + (cae_h - qr_sz) / 2

    try:
        qr_buf = _qr_image(cuit, pto_venta, tipo_cbte, nro_cbte,
                           fecha_raw, imp_total, doc_tipo, doc_nro, cae)
        c.drawImage(ImageReader(qr_buf), qr_x, qr_y, qr_sz, qr_sz,
                    preserveAspectRatio=True)
    except Exception:
        txt(qr_x + qr_sz/2, qr_y + qr_sz/2, '[QR no disponible]',
            size=7, color=GRIS, align='center')

    # texto CAE (izquierda)
    xl3   = ML + 0.4 * cm
    y_cae = y_cae_t - 0.7 * cm

    txt(xl3, y_cae,             'CÓDIGO DE AUTORIZACIÓN ELECTRÓNICA (CAE)', size=8, bold=True, color=AZUL)
    txt(xl3, y_cae - 0.55*cm,  cae, size=13, bold=True)
    txt(xl3, y_cae - 1.1*cm,   f'Vencimiento CAE: {vto_disp}', size=9)

    # leyenda AFIP
    y_leg = y_cae_b + 1.2 * cm
    txt(ML + CW/4, y_leg + 0.35*cm,
        'Comprobante Autorizado', size=10, bold=True, color=VERDE, align='center')
    txt(ML + CW/4, y_leg,
        'www.afip.gob.ar', size=8, color=GRIS, align='center')

    # divisor vertical entre leyenda y QR
    xdiv = ML + CW/2
    c.setLineWidth(0.3)
    c.setStrokeColor(GRIS)
    c.line(xdiv, y_cae_b + 0.3*cm, xdiv, y_cae_t - 0.3*cm)

    # label QR
    txt(qr_x + qr_sz/2, y_cae_b + 0.25*cm,
        'Escanear para verificar en AFIP', size=7, color=GRIS, align='center')

    # =========================================================================
    # PIE DE PÁGINA
    # =========================================================================
    txt(W / 2, MB / 2,
        f'Generado por ARCA Facturación — {datetime.now().strftime("%d/%m/%Y %H:%M")}',
        size=7, color=GRIS, align='center')

    c.save()
    buf.seek(0)
    return buf
