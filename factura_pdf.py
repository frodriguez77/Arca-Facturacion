"""
Generador de PDF — Factura electrónica AFIP.
Diseño minimalista: solo líneas, Helvetica, sin colores ni cajas.
Genera dos páginas: ORIGINAL y DUPLICADO.
"""
import base64
import io
import json
import os
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
import qrcode

W, H = A4
ML = MR = 2.0 * cm
MT = 2.0 * cm
MB = 1.5 * cm
CW = W - ML - MR

# ── tablas de referencia ──────────────────────────────────────────────────────
TIPO_NOMBRE = {
    1: 'FACTURA A',         2: 'NOTA DE DÉBITO A',  3: 'NOTA DE CRÉDITO A',
    6: 'FACTURA B',         7: 'NOTA DE DÉBITO B',  8: 'NOTA DE CRÉDITO B',
    11: 'FACTURA C',        12: 'NOTA DE DÉBITO C', 13: 'NOTA DE CRÉDITO C',
}
TIPO_LETRA  = {1:'A',2:'A',3:'A', 6:'B',7:'B',8:'B', 11:'C',12:'C',13:'C'}
DOC_NOMBRE  = {80:'CUIT', 86:'CUIL', 96:'DNI', 99:'Consumidor Final'}
CONCEPTO_NOM= {1:'Productos', 2:'Honorarios profesionales', 3:'Productos y Servicios'}
IVA_EMISOR  = {'A':'Responsable Inscripto', 'B':'Responsable Inscripto', 'C':'Monotributista'}

# ── número a letras (español) ─────────────────────────────────────────────────
_U = ['','UN','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
      'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS',
      'DIECISIETE','DIECIOCHO','DIECINUEVE']
_V = {1:'VEINTIUNO',2:'VEINTIDÓS',3:'VEINTITRÉS',4:'VEINTICUATRO',
      5:'VEINTICINCO',6:'VEINTISÉIS',7:'VEINTISIETE',8:'VEINTIOCHO',9:'VEINTINUEVE'}
_D = ['','','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA']
_C = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
      'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']

def _words(n):
    n = int(n)
    if n == 0:   return 'CERO'
    if n < 20:   return _U[n]
    if n < 30:   return 'VEINTE' if n == 20 else _V[n-20]
    if n < 100:
        t,u = divmod(n,10)
        return _D[t] + (' Y '+_U[u] if u else '')
    if n == 100: return 'CIEN'
    if n < 1000:
        h,r = divmod(n,100)
        return _C[h] + (' '+_words(r) if r else '')
    if n < 1_000_000:
        th,r = divmod(n,1000)
        pre  = 'MIL' if th==1 else _words(th)+' MIL'
        return pre + (' '+_words(r) if r else '')
    m,r = divmod(n,1_000_000)
    pre = 'UN MILLÓN' if m==1 else _words(m)+' MILLONES'
    return pre + (' '+_words(r) if r else '')

def _son(total):
    v = round(float(total), 2)
    e = int(v)
    d = round((v - e) * 100)
    return f'SON PESOS: {_words(e)} CON {d:02d}/100'

# ── helpers de formato ────────────────────────────────────────────────────────
def _fmt(v):
    v = round(float(v or 0), 2)
    neg = v < 0
    av  = abs(v)
    i,d = int(av), round((av % 1)*100)
    s = f"{i:,}".replace(',','.') + f',{d:02d}'
    return ('-' if neg else '') + s

def _cuit(c):
    c = str(c)
    return f'{c[:2]}-{c[2:10]}-{c[10]}' if len(c)==11 else c

def _parse_fecha(raw):
    for fmt in ('%Y-%m-%d','%Y%m%d','%Y-%m-%d %H:%M:%S'):
        try: return datetime.strptime(str(raw).strip()[:len(fmt)], fmt)
        except: pass
    return None

def _dfmt(raw):
    dt = _parse_fecha(raw)
    return dt.strftime('%d/%m/%Y') if dt else str(raw)[:10]

# ── QR AFIP ───────────────────────────────────────────────────────────────────
def _qr(cuit, pv, tipo, nro, fecha_raw, total, doc_tipo, doc_nro, cae):
    dt  = _parse_fecha(fecha_raw)
    p   = {"ver":1,"fecha":dt.strftime('%Y-%m-%d') if dt else str(fecha_raw)[:10],
           "cuit":int(cuit),"ptoVta":int(pv),"tipoCmp":int(tipo),"nroCmp":int(nro),
           "importe":round(float(total),2),"moneda":"PES","ctz":1,
           "tipoDocRec":int(doc_tipo),"nroDocRec":int(str(doc_nro).split('.')[0] or 0),
           "tipoCodAut":"E","codAut":int(cae)}
    b64 = base64.b64encode(json.dumps(p,separators=(',',':')).encode()).decode()
    url = f"https://www.afip.gob.ar/fe/qr/?p={b64}"
    q   = qrcode.QRCode(version=1, box_size=5, border=2,
                        error_correction=qrcode.constants.ERROR_CORRECT_M)
    q.add_data(url); q.make(fit=True)
    img = q.make_image(fill_color='black', back_color='white')
    buf = io.BytesIO(); img.save(buf, format='PNG'); buf.seek(0)
    return buf

# ── página ────────────────────────────────────────────────────────────────────
def _page(c, titulo, emp, reg, res, cliente=None, forma_pago=''):

    # extraer datos
    cuit      = str(emp['cuit'])
    nom_emp   = emp.get('nombre','')
    domicilio = emp.get('domicilio','')
    telefono  = emp.get('telefono','')
    localidad = emp.get('localidad','')
    ing_brutos= emp.get('ing_brutos','')
    inicio    = emp.get('inicio_actividades','')
    homo      = emp.get('homologacion', False)
    matricula = emp.get('matricula', '')
    logo_path = emp.get('logo_path', '')

    tipo_cbte = int(reg['tipo_cbte'])
    pto_venta = int(reg['punto_venta'])
    concepto  = int(reg['concepto'])
    doc_tipo  = int(reg['doc_tipo'])
    doc_nro   = str(reg['doc_nro']).split('.')[0]
    razon_soc = str(reg['razon_social'])
    fecha_raw = str(reg['fecha'])
    imp_neto  = float(reg.get('imp_neto', 0))
    alicuota  = float(reg.get('alicuota', 0))
    imp_iva   = float(reg.get('imp_iva',  0))
    imp_total = float(reg.get('imp_total',0))
    descrip   = str(reg.get('descripcion', CONCEPTO_NOM.get(concepto,'')))

    nro_cbte  = int(res['nro'])
    cae       = str(res['cae'])
    vto_cae   = str(res.get('vto_cae',''))
    cliente_domicilio = (cliente or {}).get('domicilio', '')

    letra    = TIPO_LETRA.get(tipo_cbte,'?')
    tipo_nom = TIPO_NOMBRE.get(tipo_cbte, f'Tipo {tipo_cbte}')
    nro_fmt  = f'{pto_venta:05d}-{nro_cbte:08d}'
    doc_nom  = DOC_NOMBRE.get(doc_tipo, f'Doc.{doc_tipo}')
    conc_nom = CONCEPTO_NOM.get(concepto, str(concepto))

    # helpers
    NEGRO = colors.black
    GRIS  = colors.HexColor('#888888')

    def hline(y, x1=ML, x2=ML+CW, lw=0.4, color=NEGRO):
        c.setStrokeColor(color); c.setLineWidth(lw)
        c.line(x1, y, x2, y)

    def T(x, y, s, sz=9, bold=False, align='left', color=NEGRO, font='Helvetica'):
        c.setFillColor(color)
        c.setFont(font+('-Bold' if bold else ''), sz)
        s = str(s)
        {'right': c.drawRightString, 'center': c.drawCentredString}.get(
            align, c.drawString)(x, y, s)
        c.setFillColor(NEGRO)

    # ── ENCABEZADO ──────────────────────────────────────────────────
    y = H - MT

    # cuadro de letra — derecha arriba (3×3 cm)
    box_sz = 2.8*cm
    bx = ML + CW - box_sz
    by = y - box_sz
    c.setStrokeColor(NEGRO); c.setLineWidth(0.8)
    c.rect(bx, by, box_sz, box_sz, fill=0, stroke=1)
    T(bx + box_sz/2, by + box_sz/2 - 0.55*cm, letra, sz=48, bold=True, align='center')
    T(bx + box_sz/2, by + 0.25*cm, f'Código {tipo_cbte:02d}', sz=7, align='center', color=GRIS)

    # tipo + número + fecha + titulo — a la izquierda del cuadro
    xr = bx - 0.4*cm
    T(xr, y - 0.65*cm, tipo_nom,            sz=10, bold=True, align='right')
    T(xr, y - 1.15*cm, f'Nro  {nro_fmt}',  sz=9,  align='right')
    T(xr, y - 1.60*cm, f'Fecha  {_dfmt(fecha_raw)}', sz=9, align='right')
    T(xr, y - 2.05*cm, titulo,              sz=9,  bold=True, align='right', color=GRIS)

    # logo — izquierda (si existe)
    x_emp = ML
    if logo_path and os.path.exists(logo_path):
        try:
            logo_max_w = 3.2 * cm
            logo_max_h = box_sz * 0.80
            c.drawImage(ImageReader(logo_path), ML, y - logo_max_h,
                        logo_max_w, logo_max_h, preserveAspectRatio=True, mask='auto')
            x_emp = ML + logo_max_w + 0.4 * cm
        except Exception:
            pass

    # nombre empresa — grande izquierda (o a la derecha del logo)
    T(x_emp, y - 0.7*cm, nom_emp, sz=15 if x_emp > ML else 16, bold=True)

    # datos fiscales empresa — bajo el nombre
    ye = y - 1.15*cm
    T(x_emp, ye,             _cuit(cuit),               sz=8, color=GRIS)
    T(x_emp, ye - 0.42*cm,   IVA_EMISOR.get(letra,''),  sz=8, color=GRIS)
    info = '  ·  '.join(filter(None,[domicilio, telefono, localidad]))
    if info: T(x_emp, ye - 0.84*cm, info, sz=8, color=GRIS)
    if ing_brutos: T(x_emp, ye - 1.26*cm, f'Ing. Brutos: {ing_brutos}', sz=8, color=GRIS)
    if inicio:     T(x_emp, ye - (1.68 if ing_brutos else 1.26)*cm,
                     f'Inicio actividades: {inicio}', sz=8, color=GRIS)
    if homo:       T(x_emp, ye - 2.1*cm, 'HOMOLOGACIÓN', sz=8, bold=True, color=colors.red)

    # línea separadora gruesa bajo encabezado
    y_sep1 = y - box_sz - 0.45*cm
    hline(y_sep1, lw=0.8)

    # ── RECEPTOR ────────────────────────────────────────────────────
    yr = y_sep1 - 0.5*cm
    pad = 2.8*cm   # ancho columna etiqueta

    def fila_rec(label, valor, dy):
        T(ML,        yr-dy, label, sz=8, color=GRIS)
        T(ML+pad,    yr-dy, valor, sz=9)

    fila_rec('Señor/es',       razon_soc,  0.00*cm)
    fila_rec(doc_nom,          doc_nro,    0.55*cm)
    fila_rec('Condición IVA',  'Responsable monotributo' if letra=='C' else 'Responsable Inscripto',
             1.10*cm)
    fila_rec('Concepto',       conc_nom,   1.65*cm)

    y_rec_extra = 0
    if cliente_domicilio:
        T(ML,      yr - 2.20*cm, 'Domicilio',            sz=8, color=GRIS)
        T(ML+pad,  yr - 2.20*cm, cliente_domicilio[:80], sz=8)
        y_rec_extra = 0.55*cm

    fp = (forma_pago or '').strip() or 'Contado'
    T(ML,     yr - 2.20*cm - y_rec_extra, 'Cond. Venta', sz=8, color=GRIS)
    T(ML+pad, yr - 2.20*cm - y_rec_extra, fp,            sz=8)
    y_rec_extra += 0.55*cm

    y_sep2 = yr - 2.20*cm - y_rec_extra
    hline(y_sep2)

    # ── ÍTEMS ───────────────────────────────────────────────────────
    yi = y_sep2 - 0.45*cm
    T(ML,        yi, 'Descripción', sz=8, color=GRIS)
    T(ML+CW,     yi, 'Importe',     sz=8, color=GRIS, align='right')
    hline(yi - 0.25*cm, lw=0.3, color=GRIS)

    yi2 = yi - 0.70*cm
    T(ML,    yi2, descrip,       sz=9)
    T(ML+CW, yi2, _fmt(imp_neto), sz=9, align='right')

    y_sep3 = yi2 - 1.8*cm
    hline(y_sep3)

    # ── TOTALES ─────────────────────────────────────────────────────
    xt  = ML + CW * 0.55
    xt2 = ML + CW

    def fila_tot(label, valor, dy, bold=False):
        T(xt,  y_sep3-dy, label, sz=9 if not bold else 10,
          bold=bold, align='right', color=NEGRO if bold else GRIS)
        T(xt2, y_sep3-dy, valor, sz=9 if not bold else 10,
          bold=bold, align='right')

    fila_tot('SubTotal',  _fmt(imp_neto),  0.50*cm)
    if imp_iva > 0:
        iva_lbl = f'IVA {alicuota:.0f}%' if alicuota else 'IVA'
        fila_tot(iva_lbl, _fmt(imp_iva),   1.00*cm)

    hline(y_sep3 - (1.35 if imp_iva > 0 else 0.85)*cm, x1=xt-0.2*cm, lw=0.4)
    fila_tot('TOTAL',     _fmt(imp_total), (1.75 if imp_iva > 0 else 1.25)*cm, bold=True)

    y_sep4 = y_sep3 - (2.20 if imp_iva > 0 else 1.70)*cm
    hline(y_sep4)

    # ── SON MONEDA ──────────────────────────────────────────────────
    T(ML, y_sep4 - 0.50*cm, _son(imp_total), sz=8, color=GRIS)

    y_sep5 = y_sep4 - 0.90*cm
    hline(y_sep5, lw=0.3, color=GRIS)

    # ── MATRÍCULA ───────────────────────────────────────────────────
    mat_extra = 0
    if matricula:
        mat_h  = 0.85 * cm
        mat_y0 = y_sep5 - 0.20 * cm - mat_h
        c.setStrokeColor(NEGRO); c.setLineWidth(0.5)
        c.rect(ML, mat_y0, CW, mat_h, fill=0, stroke=1)
        T(ML + CW / 2, mat_y0 + mat_h / 2 - 0.15 * cm,
          matricula, sz=9, align='center', font='Helvetica-Oblique')
        mat_extra = mat_h + 0.35 * cm

    # ── QR + CAE ────────────────────────────────────────────────────
    qr_sz = 3.0*cm
    y_qr  = y_sep5 - mat_extra
    qr_y  = y_qr - qr_sz - 0.40*cm

    try:
        qr_buf = _qr(cuit, pto_venta, tipo_cbte, nro_cbte,
                     fecha_raw, imp_total, doc_tipo, doc_nro, cae)
        c.drawImage(ImageReader(qr_buf), ML, qr_y, qr_sz, qr_sz,
                    preserveAspectRatio=True)
    except Exception:
        pass

    xc = ML + qr_sz + 0.5*cm
    T(xc, y_qr - 0.60*cm, f'CAE: {cae}',              sz=9, bold=True)
    T(xc, y_qr - 1.05*cm, f'Vto. CAE: {_dfmt(vto_cae)}', sz=9)
    T(xc, y_qr - 1.50*cm, 'Comprobante Autorizado',    sz=8, color=GRIS)
    T(xc, y_qr - 1.90*cm, 'www.afip.gob.ar',           sz=8, color=GRIS)

    # ── PIE ─────────────────────────────────────────────────────────
    hline(MB + 0.6*cm, lw=0.3, color=GRIS)
    now = datetime.now()
    pie = f"{now.strftime('%b')} {now.day} {now.strftime('%Y  %I:%M %p')}  —  Generado por ARCA Facturación"
    T(W/2, MB + 0.20*cm, pie, sz=7, align='center', color=GRIS)


# ── entrada pública ───────────────────────────────────────────────────────────
def generar_pdf(empresa, registro, resultado, cliente=None, forma_pago=''):
    """
    Retorna BytesIO con el PDF (ORIGINAL + DUPLICADO).
    empresa    : dict {nombre, cuit, domicilio, telefono, localidad,
                       ing_brutos, inicio_actividades, homologacion}
    registro   : dict {punto_venta, tipo_cbte, concepto, doc_tipo, doc_nro,
                       razon_social, fecha, imp_neto, alicuota, imp_iva,
                       imp_total, descripcion (opcional)}
    resultado  : dict {nro, cae, vto_cae}
    cliente    : dict opcional {domicilio, nombre, estado}
    forma_pago : str opcional — condición de venta (default "Contado")
    """
    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    _page(c, 'ORIGINAL',  empresa, registro, resultado, cliente, forma_pago)
    c.showPage()
    _page(c, 'DUPLICADO', empresa, registro, resultado, cliente, forma_pago)
    c.save()
    buf.seek(0)
    return buf
