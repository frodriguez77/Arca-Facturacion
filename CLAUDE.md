# CLAUDE.md — Manual de desarrollo para Arca-Facturacion

## Regla principal: siempre AGREGAR, nunca romper

Antes de editar cualquier archivo, respondé estas preguntas:
1. ¿Este cambio toca funcionalidad que ya existe y funciona?  
2. ¿Si lo toca, estoy preservando exactamente el comportamiento actual?  
3. ¿El usuario me pidió **explícitamente** modificar esa parte?

Si la respuesta a 3 es "no", **no la toques**. Agregá código nuevo al lado, no sobre lo existente.

---

## Arquitectura general

| Componente | Descripción |
|---|---|
| `app.py` | Flask — todas las rutas API y vistas |
| `factura_pdf.py` | Generación de PDF con ReportLab |
| `wsaa.py` | Autenticación AFIP (TA/token) |
| `wsfe.py` | Web Service Factura Electrónica AFIP |
| `repository.py` | Lectura/escritura de `empresas.json` y `usuarios.json` |
| `config.py` | URLs de AFIP (homo/prod) |
| `openssl_util.py` | Localiza OpenSSL en el sistema |
| `templates/` | Jinja2 HTML — index, admin, reportes, login, error |
| `static/logos/` | Logos de empresa subidos (no en git) |
| `uploads/{CUIT}/{YYYY-MM}/` | Excel de entrada y resultado por empresa y mes |
| `certificados/{CUIT}/` | Claves `.key` y `.crt` de AFIP |

---

## Archivos de datos — NUNCA en git, NUNCA sobreescribir

| Archivo | Contenido | Regla |
|---|---|---|
| `empresas.json` | Config de cada empresa | Solo Repository.update(), nunca escribir directo |
| `usuarios.json` | Usuarios y contraseñas (hash) | Solo Repository |
| `email_config.json` | SMTP global (creado en runtime) | Solo a través de `/api/config-email` |
| `uploads/**/*.xlsx` | Excel del cliente con datos a facturar | **Solo append, NUNCA sobreescribir** |
| `uploads/**/facturas_resultado.xlsx` | Resultados de AFIP | **Solo append de filas nuevas** |
| `certificados/**` | Claves privadas AFIP | No tocar desde el código |

---

## Tipos de comprobante AFIP

```python
_TIPOS_FACTURA = {1, 6, 11, 51, 201, 206, 211}   # Facturas A/B/C/M/FCE
_TIPOS_NC      = {3, 8, 13, 53, 203, 208, 213}    # Notas de Crédito
_TIPOS_ND      = {2, 7, 12, 52, 202, 207, 212}    # Notas de Débito
```

Función helper: `_tipo_grupo(tipo_cbte) → 'factura' | 'nc' | 'nd'`

---

## Rutas API — mapa completo

### Empresas
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/api/empresas` | Lista empresas del usuario |
| POST | `/api/empresas` | Agrega empresa (admin) |
| PUT | `/api/empresas/<id>` | Edita empresa (admin) |
| DELETE | `/api/empresas/<id>` | Elimina empresa (admin) |
| POST | `/api/upload-logo` | Sube logo → `static/logos/` |
| GET | `/api/logo-preview` | Sirve logo seguro (path dentro de BASE) |

### Facturación
| Método | Ruta | Descripción |
|---|---|---|
| POST | `/upload` | Recibe Excel, preview, NO factura aún |
| POST | `/procesar` | Envía a AFIP, escribe resultado (append-only) |
| GET | `/api/resultados-mes` | Resultados de un mes (auto-detecta último si no se pide mes) |
| GET | `/api/meses-disponibles` | Lista meses con `facturas_resultado.xlsx` para una empresa |
| POST | `/api/nota-credito` | Emite NC asociada a una factura aprobada |
| POST | `/api/nota-debito` | Emite ND asociada a una factura aprobada |
| GET | `/descargar` | Descarga Excel de resultados del mes |
| GET | `/pdf/<empresa_id>/<fila>` | Genera PDF de un comprobante |
| POST | `/imprimir` | PDF desde JSON directo |
| GET | `/plantilla` | Descarga plantilla Excel vacía |

### Correo
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/api/config-email` | Lee config SMTP global (devuelve password enmascarada) |
| POST | `/api/config-email` | Guarda config SMTP global en `email_config.json` |
| POST | `/api/enviar-factura` | Envía PDF por email — usa SMTP de la empresa si tiene, sino global |

### Usuarios (admin)
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/api/usuarios` | Lista usuarios |
| POST | `/api/usuarios` | Agrega usuario |
| PUT | `/api/usuarios/<uid>` | Edita usuario |
| DELETE | `/api/usuarios/<uid>` | Elimina usuario |
| POST | `/api/usuarios/<uid>/cambiar-password` | Cambia contraseña |

### Reportes
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/reportes` | Vista de reportes |
| GET | `/api/reportes` | Datos de reportes (acepta `?tipo=`, `?desde=`, `?hasta=`) |
| GET | `/api/reportes/exportar` | Exporta reporte a Excel |

### Admin
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/admin` | Panel de administración |
| GET | `/admin/backup` | Descarga ZIP con empresas.json + usuarios.json |
| POST | `/admin/restore` | Restaura backup |
| POST | `/generar-csr` | Genera clave .key y .csr para AFIP |

---

## Modelo de datos de empresa

Campos guardados en `empresas.json` por cada empresa:

```json
{
  "id": "string",
  "nombre": "string",
  "cuit": "11 dígitos",
  "cert": "ruta al .crt",
  "key": "ruta al .key",
  "homologacion": false,
  "domicilio": "",
  "telefono": "",
  "localidad": "",
  "ing_brutos": "",
  "inicio_actividades": "",
  "matricula": "",
  "logo_path": "ruta absoluta o vacío",
  "smtp_server": "",
  "smtp_port": 587,
  "smtp_user": "",
  "smtp_password": "",
  "smtp_ssl": false,
  "nombre_remitente": ""
}
```

Al agregar un campo nuevo a la empresa:
1. Agregar en `api_empresa_add` (POST)
2. Agregar en `api_empresa_edit` (PUT) — si es contraseña: solo actualizar si no viene enmascarada
3. Agregar el input en el form de empresa en `admin.html`
4. Poblar en `editarEmpresa()` del JS
5. Limpiar en `limpiarForm()` del JS
6. Incluir en el `body` del submit del form

---

## Invariantes críticos — NUNCA violar

- **Excel: append-only.** Los archivos de resultados solo reciben filas nuevas. Nunca `df.to_excel()` sin primero verificar que es una fila nueva.
- **CUIT en rutas.** La carpeta de cada empresa es `uploads/{CUIT}/`. Si el CUIT cambia, los archivos históricos quedan en la carpeta vieja (comportamiento aceptado).
- **`_tipo_grupo()`** debe cubrir todos los tipos. Si AFIP agrega tipos nuevos, agregar a los sets `_TIPOS_*` y a `TIPO_NOMBRE`.
- **Contraseña SMTP enmascarada.** Si el campo llega con `•`, no sobreescribir la contraseña existente.
- **Seguridad en logo-preview.** Siempre verificar `os.path.realpath(path).startswith(BASE)` antes de servir un archivo.
- **`_user_can_access(user, empresa_id)`** debe chequearse en TODOS los endpoints que operan sobre datos de una empresa.

---

## Flujo de facturación (secuencia)

```
1. Usuario selecciona empresa  →  card-empresa
2. Upload Excel  →  POST /upload  →  preview en card-preview
3. Confirmar  →  POST /procesar  →  llama a AFIP  →  guarda resultado  →  card-result
4. Desde card-result: PDF / Email / NC / ND / Descargar Excel
5. "Nueva facturación"  →  vuelve a card-upload (con btn "Ver resultados" visible)
6. Selector de mes en card-result  →  GET /api/resultados-mes?mes=YYYY-MM
```

---

## SMTP — Prioridad de configuración

1. **Por empresa** (`empresas.json` → campos `smtp_*`): tiene prioridad.
2. **Global** (`email_config.json` vía `/api/config-email`): fallback cuando la empresa no tiene SMTP propio.

Lógica en `api_enviar_factura`:
```python
if empresa.get('smtp_server') and empresa.get('smtp_user') and empresa.get('smtp_password'):
    cfg = {campos de la empresa}
else:
    cfg = _load_email_config()
    if not completo: error
```

---

## Estructura de `templates/index.html` — cards y navegación

| Card ID | Cuándo se muestra |
|---|---|
| `card-empresa` | Al entrar / al cambiar empresa |
| `card-upload` | Después de elegir empresa / "Nueva facturación" |
| `card-preview` | Después de subir Excel exitosamente |
| `card-result` | Después de enviar a ARCA / al cargar resultados del mes |

Funciones JS clave:
- `cargarResultadosMes(mes='')` — carga resultados, actualiza selector de mes
- `cargarSelectorMeses()` — puebla `<select id="sel-mes">` con meses disponibles
- `renderResultados(data)` — dibuja la tabla de resultados
- `show(id)` / `hide(id)` — mostrar/ocultar cards

---

## Branch y despliegue

- Branch de desarrollo: `claude/new-pc-download-setup-5AF1K`
- Repo: `frodriguez77/Arca-Facturacion`
- PC del usuario: `D:\Arca-Facturacion`
- Script de actualización en PC: `actualizar.ps1` (PowerShell) — checkea archivos específicos del branch
- Archivos actualizados por el script: `app.py`, `factura_pdf.py`, `wsfe.py`, `wsaa.py`, `templates/admin.html`, `templates/index.html`, `templates/reportes.html`
- **Los archivos de datos (`empresas.json`, `usuarios.json`, `email_config.json`, `uploads/`, `certificados/`) NO están en git y NO son tocados por el script.**

---

## Versionado

El archivo `VERSION` (en la raíz) contiene la versión actual en formato `MAJOR.MINOR.PATCH`.

```
1.0.0   ← MAJOR.MINOR.PATCH
│ │ └── Corrección de bug (sin features nuevas)
│ └──── Feature nueva o mejora (no rompe nada existente)
└────── Cambio estructural grande (rediseño, migración de datos)
```

**Cuándo subir la versión:**

| Tipo de cambio | Qué hacer |
|---|---|
| Solo bugs corregidos | Incrementar PATCH: `1.0.0 → 1.0.1` |
| Feature nueva / mejora visible | Incrementar MINOR: `1.0.0 → 1.1.0` |
| Cambio estructural / breaking | Incrementar MAJOR: `1.0.0 → 2.0.0` |

**Proceso para publicar una nueva versión:**

1. Editar el archivo `VERSION` con el número nuevo (ej. `1.1.0`)
2. Hacer commit: `git commit -am "Release v1.1.0"`
3. Crear tag: `git tag v1.1.0`
4. Push: `git push && git push --tags`
5. El actualizador en la PC (`actualizar.ps1`) descarga el archivo `VERSION` y muestra el número al terminar.

**La versión aparece en:**
- Footer de cada pantalla (index, admin, reportes)
- Mensaje de cierre de `actualizar.ps1`
- Variable de Flask: `APP_VERSION` (inyectada vía context processor en todos los templates como `{{ app_version }}`)

---

## Checklist antes de hacer un PR / push

- [ ] ¿Rompí alguna funcionalidad existente?
- [ ] ¿Agregué campos a empresa? → Seguí los 6 pasos del modelo de datos
- [ ] ¿Nuevo endpoint? → Agregarlo al mapa de rutas en este archivo
- [ ] ¿Toco Excel? → Solo append
- [ ] ¿Toco logo/archivo? → Verificar path dentro de BASE
- [ ] ¿Toco autenticación? → `_user_can_access` en todos los endpoints
- [ ] ¿Python válido? → `python3 -c "import ast; ast.parse(open('app.py').read())"`
- [ ] ¿Es una nueva versión? → Actualizar `VERSION` y crear tag git
