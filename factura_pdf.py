"""
Generador de PDF — Factura electrónica AFIP.
Diseño: encabezado 3 columnas, receptor, cuerpo, pie con QR.
Genera dos páginas: ORIGINAL y DUPLICADO.
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

W, H = A4          # 595.28 x 841.89 pts
ML = MR = 1.5 * cm
MT = MB = 1.5 * cm
CW = W - ML - MR   # ancho útil

# ── tablas de referencia ──────────────────────────────────────────────────────
TIPO_NOMBRE = {
    1: 'FACTURA A',         2: 'NOTA DE DÉBITO A',  3: 'NOTA DE CRÉDITO A',
    6: 'FACTURA B',         7: 'NOTA DE DÉBITO B',  8: 'NOTA DE CRÉDITO B',
    11: 'FACTURA C',        12: 'NOTA DE DÉBITO C', 13: 'NOTA DE CRÉDITO C',
}
TIPO_LETRA = {1:'A',2:'A',3:'A', 6:'B',7:'B',8:'B', 11:'C',12:'C',13:'C'}
DOC_NOMBRE = {80:'CUIT', 86:'CUIL', 96:'DNI', 99:'Consumidor Final'}
CONCEPTO_NOMBRE = {1:'Productos', 2:'Servicios', 3:'Productos y Servicios'}
IVA_EMISOR = {
    'A': 'IVA RESPONSABLE INSCRIPTO',
    'B': 'IVA RESPONSABLE INSCRIPTO',
    'C': 'MONOTRIBUTISTA',
}

# ── número a letras (español) ────────────────────────────────────────────────
_UNID = [
    '', 'UN', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE',
    'DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS',
    'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE',
]
_VEINT = {
    1:'VEINTIUNO',2:'VEINTIDÓS',3:'VEINTITRÉS',4:'VEINTICUATRO',
    5:'VEINTICINCO',6:'VEINTISÉIS',7:'VEINTISIETE',8:'VEINTIOCHO',9:'VEINTINUEVE',
}
_DECE  = ['','','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA']
_CENT  = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
          'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']

def _words(n):
    n = int(n)
    if n == 0:    return 'CERO'
    if n < 20:    return _UNID[n]
    if n < 30:    return 'VEINTE' if n == 20 else _VEINT[n - 20]
    if n < 100:
        t, u = divmod(n, 10)
        return _DECE[t] + (' Y ' + _UNID[u] if u else '')
    if n == 100:  return 'CIEN'
    if n < 1000:
        h, r = divmod(n, 100)
        return _CENT[h] + (' ' + _words(r) if r else '')
    if n < 1000000:
        th, r = divmod(n, 1000)
        pre = 'MIL' if th == 1 else _words(th) + ' MIL'
        return pre + (' ' + _words(r) if r else '')
    m, r = divmod(n, 1000000)
    pre = 'UN MILLÓN' if m == 1 else _words(m) + ' MILLONES'
    return pre + (' ' + _words(r) if r else '')

def _son_moneda(total):
    v      = round(float(total), 2)
    entero = int(v)
    cents  = round((v - entero) * 100)
    return f'SON PESOS: {_words(entero)} CON {cents:02d}/100'


# ── helpers de formato ────────────────────────────────────────────────────────
def _fmt(v):
    v = round(float(v or 0), 2)
    i, d = int(abs(v)), round((abs(v) % 1) * 100)
    s = f"{i:,}".replace(",", ".") + f",{d:02d}"
    return ('-' if v < 0 else '') + s

def _cuit_fmt(c):
    c = str(c)
    return f"{c[:2]}-{c[2:10]}-{c[10]}" if len(c) == 11 else c

def _parse_fecha(raw):
    for fmt in ('%Y-%m-%d', '%Y%m%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(str(raw).strip()[:len(fmt)], fmt)
        except Exception:
            pass
    return None

def _dfmt(raw):
    dt = _parse_fecha(raw)
    return dt.strftime('%d/%m/%Y') if dt else str(raw)[:10]


# ── QR AFIP ───────────────────────────────────────────────────────────────────
def _qr_buf(cuit, pv, tipo, nro, fecha_raw, total, doc_tipo, doc_nro, cae):
    dt       = _parse_fecha(fecha_raw)
    fecha    = dt.strftime('%Y-%m-%d') if dt else str(fecha_raw)[:10]
    payload  = {
        "ver":1,"fecha":fecha,"cuit":int(cuit),"ptoVta":int(pv),
        "tipoCmp":int(tipo),"nroCmp":int(nro),"importe":round(float(total),2),
        "moneda":"PES","ctz":1,"tipoDocRec":int(doc_tipo),
        "nroDocRec":int(str(doc_nro).split('.')[0] or 0),
        "tipoCodAut":"E","codAut":int(cae),
    }
    b64 = base64.b64encode(json.dumps(payload, separators=(',',':')).encode()).decode()
    url = f"https://www.afip.gob.ar/fe/qr/?p={b64}"
    qr  = qrcode.QRCode(version=1, box_size=5, border=2,
                        error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


# ── dibujo de una página ──────────────────────────────────────────────────────
def _draw_page(c, titulo, empresa, registro, resultado):

    # ── extraer datos ──────────────────────────────────────────────
    cuit      = str(empresa['cuit'])
    nom_emp   = empresa.get('nombre', '')
    domicilio = empresa.get('domicilio', '')
    telefono  = empresa.get('telefono', '')
    localidad = empresa.get('localidad', '')
    ing_brutos= empresa.get('ing_brutos', '')
    inicio_act= empresa.get('inicio_actividades', '')
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
    descrip   = str(registro.get('descripcion', CONCEPTO_NOMBRE.get(concepto, '')))

    nro_cbte  = int(resultado['nro'])
    cae       = str(resultado['cae'])
    vto_cae   = str(resultado.get('vto_cae', ''))

    letra     = TIPO_LETRA.get(tipo_cbte, '?')
    tipo_nom  = TIPO_NOMBRE.get(tipo_cbte, f'Tipo {tipo_cbte}')
    nro_fmt   = f"{tipo_nom} {letra}{pto_venta:05d}-{nro_cbte:08d}"
    doc_nom   = DOC_NOMBRE.get(doc_tipo, f'Doc.{doc_tipo}')
    iva_emis  = IVA_EMISOR.get(letra, '')

    # ── helpers de canvas ──────────────────────────────────────────
    NEGRO = colors.black
    GRIS  = colors.HexColor('#555555')

    def hline(y, x1=ML, x2=ML+CW, lw=0.5):
        c.setStrokeColor(NEGRO); c.setLineWidth(lw)
        c.line(x1, y, x2, y)

    def vline(x, y1, y2, lw=0.5):
        c.setStrokeColor(NEGRO); c.setLineWidth(lw)
        c.line(x, y1, x, y2)

    def txt(x, y, s, size=8, bold=False, font='Courier', align='left', color=NEGRO):
        c.setFillColor(color)
        fname = (font + '-Bold') if bold else font
        c.setFont(fname, size)
        s = str(s)
        if   align == 'right':  c.drawRightString(x, y, s)
        elif align == 'center': c.drawCentredString(x, y, s)
        else:                   c.drawString(x, y, s)
        c.setFillColor(NEGRO)

    # ── coordenadas Y de cada sección (desde abajo, en pts) ────────
    y_bot   = MB                      # 1.5 cm
    y_tot   = y_bot   + 1.4 * cm     # barra SubTotal/IVA/TOTAL
    y_qr    = y_tot   + 4.2 * cm     # bloque QR + CAE
    y_son   = y_qr    + 1.3 * cm     # SON Moneda
    y_obs   = y_son   + 1.4 * cm     # Observaciones
    y_leg   = y_obs   + 2.8 * cm     # texto legal
    y_items = y_leg   + 6.0 * cm     # cuerpo ítems
    y_ihdr  = y_items + 0.75* cm     # cabecera de ítems
    y_rec   = y_ihdr  + 3.7 * cm     # sección receptor
    y_top   = H - MT                  # tope del contenido

    # ── borde exterior ─────────────────────────────────────────────
    c.setStrokeColor(NEGRO); c.setLineWidth(0.8)
    c.rect(ML, y_bot, CW, y_top - y_bot, fill=0, stroke=1)

    # ══════════════════════════════════════════════════════════════
    # ENCABEZADO
    # ══════════════════════════════════════════════════════════════
    hline(y_rec)      # línea inferior del encabezado

    # columnas del encabezado
    col_lw = ML + CW * 0.42   # divisor izq  (42 % desde izq)
    col_rw = ML + CW * 0.58   # divisor der  (58 % desde izq)
    vline(col_lw, y_rec, y_top)
    vline(col_rw, y_rec, y_top)

    # -- columna IZQUIERDA: datos empresa --
    xl = ML + 0.35 * cm
    yt = y_top - 0.55 * cm
    txt(xl, yt,             nom_emp,  size=10, bold=True)
    if domicilio:  txt(xl, yt - 0.50*cm, domicilio, size=8)
    if telefono:   txt(xl, yt - 0.95*cm, f'Tel / Fax: {telefono}', size=8)
    if localidad:  txt(xl, yt - 1.40*cm, localidad, size=8)
    txt(xl, yt - (1.85 if localidad else 1.40)*cm, iva_emis, size=8, bold=True)
    if homo:
        txt(xl, yt - 2.35*cm, 'HOMOLOGACIÓN', size=7, color=colors.red, bold=True)

    # -- columna CENTRO: letra grande --
    xc = (col_lw + col_rw) / 2
    yc = (y_rec + y_top) / 2
    txt(xc, yc + 0.5*cm,  letra, size=52, bold=True, font='Helvetica', align='center')
    txt(xc, yc - 1.1*cm,  'Código', size=7, align='center')
    txt(xc, yc - 1.55*cm, f'{tipo_cbte:02d}', size=9, bold=True, align='center')

    # -- columna DERECHA: nro / fecha / titulo --
    xr     = col_rw + 0.35 * cm
    xr_end = ML + CW - 0.35 * cm

    txt(xr_end, y_top - 0.6*cm,  nro_fmt,                  size=9,  bold=True, align='right')
    txt(xr_end, y_top - 1.1*cm,  f'Fecha: {_dfmt(fecha_raw)}', size=9, align='right')
    txt(xr_end, y_top - 1.6*cm,  titulo,                   size=9,  bold=True, align='right')

    # datos fiscales del emisor
    yr2 = y_top - 2.3 * cm
    txt(xr, yr2,              f'CUIT: {_cuit_fmt(cuit)}',   size=8)
    if ing_brutos:
        txt(xr, yr2 - 0.45*cm, f'Ing.Brutos: {ing_brutos}', size=8)
    if inicio_act:
        txt(xr, yr2 - 0.90*cm, f'Inic. de actividades: {inicio_act}', size=8)

    # ══════════════════════════════════════════════════════════════
    # RECEPTOR
    # ══════════════════════════════════════════════════════════════
    hline(y_ihdr)    # línea inferior de receptor

    xl2  = ML + 0.6 * cm
    col2 = ML + CW * 0.62
    yr   = y_rec - 0.55 * cm

    txt(xl2,  yr,              f'Señor/es:  {razon_soc}',   size=9)
    txt(col2, yr,              f'{doc_nom}: {doc_nro}',     size=9, align='right', color=NEGRO)

    txt(xl2,  yr - 0.50*cm,   'Domicilio:',                size=9)
    txt(xl2,  yr - 1.00*cm,   'localidad:',                size=9)

    iva_rec = 'Responsable monotributo' if letra == 'C' else 'Responsable Inscripto'
    txt(xl2,  yr - 1.50*cm,   f'IVA:        {iva_rec}',    size=9)
    txt(col2, yr - 1.50*cm,   f'CUIT: {_cuit_fmt(doc_nro) if doc_tipo == 80 else ""}',
        size=9, align='right')

    txt(xl2,  yr - 2.05*cm,   'Condicion de Pago:',        size=9)

    # ══════════════════════════════════════════════════════════════
    # CABECERA DE ÍTEMS
    # ══════════════════════════════════════════════════════════════
    hline(y_items)   # línea inferior de cabecera ítems

    xi_desc = ML + 0.6 * cm
    xi_imp  = ML + CW - 0.5 * cm
    yi_hdr  = y_ihdr - 0.52 * cm

    # línea vertical precio en cabecera e ítems
    xv_imp = ML + CW * 0.82
    vline(xv_imp, y_items, y_ihdr)

    txt(xi_desc, yi_hdr, 'Descripción',   size=8, bold=True)
    txt(xi_imp,  yi_hdr, 'Importe',       size=8, bold=True, align='right')

    # ══════════════════════════════════════════════════════════════
    # CUERPO — primera línea de ítem
    # ══════════════════════════════════════════════════════════════
    yi_row = y_items - 0.55 * cm
    txt(xi_desc, yi_row, descrip,         size=9)
    txt(xi_imp,  yi_row, _fmt(imp_neto),  size=9, align='right')

    # línea vertical extendida por cuerpo
    vline(xv_imp, y_leg, y_items)

    # ══════════════════════════════════════════════════════════════
    # TEXTO LEGAL
    # ══════════════════════════════════════════════════════════════
    hline(y_leg)

    legal = (
        'El crédito fiscal discriminado en el presente comprobante, sólo podrá ser\n'
        'computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal para\n'
        'Pequeños Contribuyentes de la Ley Nº 27.618.'
    )
    yl = y_leg - 0.55 * cm
    for linea in legal.split('\n'):
        txt(ML + 0.6*cm, yl, linea, size=8)
        yl -= 0.42 * cm

    # ══════════════════════════════════════════════════════════════
    # OBSERVACIONES
    # ══════════════════════════════════════════════════════════════
    hline(y_obs)
    txt(ML + 0.6*cm, y_obs - 0.5*cm, 'Observaciones:', size=8)

    # ══════════════════════════════════════════════════════════════
    # TIMESTAMP + SON MONEDA
    # ══════════════════════════════════════════════════════════════
    hline(y_son)
    ahora = datetime.now().strftime('%b %-d %Y %I:%M%p') if hasattr(datetime, 'strftime') else ''
    try:
        ahora = datetime.now().strftime('%b %#d %Y %I:%M%p')   # Windows
    except Exception:
        try:
            ahora = datetime.now().strftime('%b %-d %Y %I:%M%p')  # Linux/Mac
        except Exception:
            ahora = datetime.now().strftime('%b %d %Y %I:%M%p')

    txt(ML + 0.6*cm, y_son - 0.45*cm, ahora,               size=8)
    txt(ML + 0.6*cm, y_son - 0.9*cm,  _son_moneda(imp_total), size=8)

    # ══════════════════════════════════════════════════════════════
    # QR + CAE
    # ══════════════════════════════════════════════════════════════
    hline(y_qr)

    qr_sz = 3.2 * cm
    qr_x  = ML + 0.4 * cm
    qr_y  = y_qr - qr_sz - (4.2*cm - qr_sz) / 2 + y_tot - y_qr + 4.2*cm / 2   # centrado vertical
    qr_y  = y_tot + (4.2*cm - qr_sz) / 2

    try:
        qr_img = _qr_buf(cuit, pto_venta, tipo_cbte, nro_cbte,
                         fecha_raw, imp_total, doc_tipo, doc_nro, cae)
        c.drawImage(ImageReader(qr_img), qr_x, qr_y, qr_sz, qr_sz,
                    preserveAspectRatio=True)
    except Exception:
        pass

    # CAE a la derecha del QR
    xc_cae = ML + 5.0 * cm
    yc_cae = y_qr - 1.2 * cm
    txt(xc_cae, yc_cae,            f'CAE: {cae}',           size=9, bold=True)
    txt(xc_cae, yc_cae - 0.55*cm,  f'VTO CAE: {_dfmt(vto_cae)}', size=9)

    # ══════════════════════════════════════════════════════════════
    # BARRA SUBTOTAL / IVA / TOTAL
    # ══════════════════════════════════════════════════════════════
    hline(y_tot)

    col_w   = CW / 3
    xsub    = ML + col_w * 0.5
    xiva    = ML + col_w * 1.5
    xtot    = ML + col_w * 2.5

    vline(ML + col_w,   y_bot, y_tot)
    vline(ML + col_w*2, y_bot, y_tot)

    y_lbl = y_tot - 0.38 * cm
    y_val = y_bot + 0.25 * cm

    txt(xsub, y_lbl, 'SubTotal', size=8, align='center')
    txt(xiva,  y_lbl, 'IVA',      size=8, align='center')
    txt(xtot,  y_lbl, 'TOTAL',    size=8, align='center')

    txt(xsub, y_val, _fmt(imp_neto),  size=9, align='center')
    txt(xiva,  y_val, _fmt(imp_iva),   size=9, align='center')
    txt(xtot,  y_val, _fmt(imp_total), size=9, bold=True, align='center')


# ── entrada pública ───────────────────────────────────────────────────────────
def generar_pdf(empresa, registro, resultado):
    """
    Retorna BytesIO con el PDF (ORIGINAL + DUPLICADO).

    empresa  : dict {nombre, cuit, domicilio, telefono, localidad,
                     ing_brutos, inicio_actividades, homologacion}
    registro : dict {punto_venta, tipo_cbte, concepto, doc_tipo, doc_nro,
                     razon_social, fecha, imp_neto, alicuota, imp_iva, imp_total,
                     descripcion (opcional)}
    resultado: dict {nro, cae, vto_cae}
    """
    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)

    _draw_page(c, 'ORIGINAL',   empresa, registro, resultado)
    c.showPage()
    _draw_page(c, 'DUPLICADO',  empresa, registro, resultado)

    c.save()
    buf.seek(0)
    return buf
