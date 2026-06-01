# ARCA - Facturación Electrónica

Sistema de facturación electrónica para AFIP (Argentina), desarrollado en Python/Flask.
Permite emitir facturas electrónicas para múltiples empresas desde cualquier PC con Windows.

---

## Requisitos previos

- **Python 3.10 o superior** → [descargar](https://www.python.org/downloads/)
  - Durante la instalación tildar **"Add Python to PATH"**
- **Git para Windows** (incluye OpenSSL) → [descargar](https://git-scm.com/download/win)
- **Google Chrome** (recomendado)

---

## Instalación en una PC nueva

### Opción A — Con Git

```bash
git clone https://github.com/frodriguez77/Arca-Facturacion.git C:\Arca-Facturacion
cd C:\Arca-Facturacion
instalar.bat
```

### Opción B — Sin Git

1. Ir a [github.com/frodriguez77/Arca-Facturacion](https://github.com/frodriguez77/Arca-Facturacion)
2. Clic en **Code → Download ZIP**
3. Descomprimir en `C:\Arca-Facturacion\`
4. Doble clic en **`instalar.bat`**

El instalador automáticamente:
- Verifica que Python esté instalado
- Instala todas las dependencias
- Crea las carpetas necesarias
- Crea el acceso directo **"ARCA Facturación"** en el escritorio

---

## Uso diario

1. Doble clic en **"ARCA Facturación"** del escritorio
2. El servidor inicia automáticamente (sin ventana)
3. Se abre Chrome en `http://localhost:5000`
4. Si el servidor ya estaba corriendo, abre Chrome directo

> **Acceso manual:** abrir Chrome y navegar a `http://localhost:5000`

---

## Primer inicio

El sistema crea automáticamente un usuario administrador:

| Usuario | Contraseña |
|---------|------------|
| `admin` | `admin123` |

⚠️ **Cambiar la contraseña** desde Admin → Usuarios → Editar.

---

## Configurar una empresa nueva

1. Ir a **Admin → pestaña 2. Generar CSR**
2. Seguir los 5 pasos del acordeón:
   - Generar la clave y el CSR
   - Subir el CSR a AFIP (Certificados Digitales)
   - Adherir el servicio de Factura Electrónica
   - Crear un punto de venta tipo **WSFEV1 Web Service**
   - Configurar la empresa en el sistema
3. Ir a **Admin → pestaña 1. Empresas → Agregar empresa**

---

## Estructura de archivos

```
C:\Arca-Facturacion\
  ├── app.py                    ← servidor Flask
  ├── repository.py             ← capa de acceso a datos
  ├── wsaa.py                   ← autenticación AFIP
  ├── wsfe.py                   ← servicio de facturación AFIP
  ├── factura_pdf.py            ← generador de PDF
  ├── config.py                 ← URLs de AFIP
  ├── openssl_util.py           ← detección de OpenSSL
  ├── instalar.bat              ← instalador (ejecutar una vez)
  ├── iniciar_servidor.bat      ← arranque del sistema
  │
  ├── certificados/             ← certificados por CUIT (NO en GitHub)
  │     └── {CUIT}/
  │           ├── {CUIT}.csr
  │           ├── {CUIT}_clave.key
  │           └── certificado.crt
  │
  ├── uploads/                  ← facturas por empresa y mes (NO en GitHub)
  │     └── {CUIT}/
  │           └── {YYYY-MM}/
  │                 ├── facturas.xlsx
  │                 └── facturas_resultado.xlsx
  │
  ├── empresas.json             ← configuración de empresas (NO en GitHub)
  └── usuarios.json             ← usuarios del sistema (NO en GitHub)
```

---

## Migrar a otra PC

### Opción recomendada (desde el sistema):

1. En la PC origen: **Admin → pestaña 4. Backup → Descargar backup completo**
2. Instalar el sistema en la PC nueva (ver sección Instalación)
3. En la PC nueva: **Admin → pestaña 4. Backup → Restaurar backup** (subir el ZIP)
4. Copiar manualmente la carpeta `certificados/` (contiene claves privadas sensibles)

### Alternativa manual:
Copiar de la PC origen a la PC nueva:
- `C:\Arca-Facturacion\uploads\`
- `C:\Arca-Facturacion\empresas.json`
- `C:\Arca-Facturacion\usuarios.json`
- `C:\Arca-Facturacion\certificados\`

---

## Formatos de factura soportados

| Tipo | Código | Descripción |
|------|--------|-------------|
| Factura A | 1 | Con IVA discriminado |
| Factura B | 6 | Consumidor final |
| Factura C | 11 | Monotributistas |
| Nota de Crédito A | 3 | |
| Nota de Crédito B | 8 | |
| Nota de Crédito C | 13 | |

---

## Plantilla Excel

Descargar desde el sistema: **botón "Plantilla Excel"** en el header.

Columnas requeridas:

| Columna | Ejemplo |
|---------|---------|
| punto_venta | 3 |
| tipo_cbte | 11 |
| concepto | 2 |
| doc_tipo | 99 |
| doc_nro | 0 |
| razon_social | Consumidor Final |
| fecha | 2026-06-01 |
| imp_neto | 1000.00 |
| alicuota | 0 |
| imp_iva | 0.00 |
| imp_total | 1000.00 |

---

## Soporte

Ante cualquier problema, revisar la consola donde corre `python app.py` — los errores se muestran con detalle completo.
