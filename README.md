# Dashboard IAAS — HGZ 1 TLAXCALA
## Guía de instalación y uso

---

## ¿Cómo funciona?

```
Tú subes el Excel a SharePoint (Office 365)
         ↓  (automático cada lunes o manual)
GitHub Actions descarga el Excel,
corre el script Python y actualiza el dashboard
         ↓
GitHub Pages publica el resultado en una URL fija
         ↓
Cualquier persona con el link lo ve actualizado
```

---

## PASO 1 — Crear la cuenta y repositorio en GitHub

1. Ve a **https://github.com** y crea una cuenta (usa un correo del hospital)
2. Haz clic en **New repository** (botón verde)
3. Configura:
   - **Repository name:** `dashboard-hgz1`
   - **Visibility:** Private ← importante
   - Marca **Add a README file**
4. Clic en **Create repository**

---

## PASO 2 — Subir los archivos a GitHub

Sube estos archivos al repositorio (arrástralos desde tu computadora):

```
dashboard-hgz1/
├── index.html                    ← copia de dashboard_completo.html (renómbralo)
├── dashboard_completo.html       ← el original (de respaldo)
├── dashboard_data.json           ← se genera automáticamente, sube uno vacío: {}
├── requirements.txt              ← incluido en este paquete
└── scripts/
    └── procesar_excel.py         ← incluido en este paquete
└── .github/
    └── workflows/
        └── actualizar_dashboard.yml  ← incluido en este paquete
```

### Para subir archivos:
1. En tu repositorio, haz clic en **Add file → Upload files**
2. Arrastra todos los archivos
3. Escribe un mensaje como "Carga inicial" y haz clic en **Commit changes**

---

## PASO 3 — Registrar la aplicación en Azure (para conectar SharePoint)

Esto le da permiso a GitHub de leer el Excel de tu SharePoint.

1. Ve a **https://portal.azure.com** e inicia sesión con tu cuenta de Office 365
2. Busca **"App registrations"** y haz clic en **New registration**
3. Configura:
   - **Name:** `Dashboard HGZ1`
   - **Supported account types:** Accounts in this organizational directory only
4. Clic en **Register**
5. Copia y guarda el **Application (client) ID** — lo necesitarás después
6. Copia y guarda el **Directory (tenant) ID**

### Crear el secreto de la app:
1. En la app que acabas de crear, ve a **Certificates & secrets**
2. Haz clic en **New client secret**
3. Descripción: `GitHub Dashboard`, Expiration: **24 months**
4. Clic en **Add**
5. **Copia el valor del secreto inmediatamente** (solo se muestra una vez)

### Dar permisos a la app:
1. Ve a **API permissions → Add a permission → Microsoft Graph**
2. Selecciona **Application permissions**
3. Busca y agrega:
   - `Sites.Read.All`
   - `Files.Read.All`
4. Haz clic en **Grant admin consent** (necesitas ser administrador)

---

## PASO 4 — Obtener el Site ID de SharePoint

1. Ve a tu sitio de SharePoint donde está el Excel
2. Abre esta URL en el navegador (reemplaza con tu dominio e hospital):
   ```
   https://graph.microsoft.com/v1.0/sites/TU-DOMINIO.sharepoint.com:/sites/TU-SITIO
   ```
   Ejemplo:
   ```
   https://graph.microsoft.com/v1.0/sites/imss.sharepoint.com:/sites/bacteriologia
   ```
3. Copia el valor del campo **"id"** de la respuesta

---

## PASO 5 — Guardar los secretos en GitHub

1. En tu repositorio de GitHub, ve a **Settings → Secrets and variables → Actions**
2. Haz clic en **New repository secret** y agrega uno por uno:

| Nombre del secret   | Valor                                              |
|---------------------|----------------------------------------------------|
| `SP_TENANT_ID`      | Directory (tenant) ID de Azure                     |
| `SP_CLIENT_ID`      | Application (client) ID de Azure                   |
| `SP_CLIENT_SECRET`  | El secreto que copiaste del paso 3                 |
| `SP_SITE_ID`        | El Site ID del paso 4                              |
| `SP_FILE_PATH`      | Ruta del Excel en SharePoint (ver ejemplo abajo)   |

### Ejemplo de SP_FILE_PATH:
```
/sites/bacteriologia/Shared Documents/Reportes IAAS/reporte_iaas.xlsx
```
Es la ruta del archivo dentro de SharePoint, comenzando con `/sites/`.

---

## PASO 6 — Activar GitHub Pages

1. En tu repositorio, ve a **Settings → Pages**
2. En **Source**, selecciona **Deploy from a branch**
3. Branch: **main**, Folder: **/ (root)**
4. Clic en **Save**
5. En unos minutos aparecerá la URL:
   ```
   https://TU-USUARIO.github.io/dashboard-hgz1/
   ```

---

## PASO 7 — Probar que todo funciona

1. Ve a tu repositorio → pestaña **Actions**
2. Haz clic en **Actualizar Dashboard HGZ 1**
3. Clic en **Run workflow → Run workflow** (botón verde)
4. Espera ~2 minutos y verifica que aparezca una palomita verde ✅
5. Abre la URL de GitHub Pages — el dashboard debe mostrar los datos del Excel

---

## Uso diario — Actualizar el dashboard

### Opción A: Automático (sin hacer nada)
El dashboard se actualiza **todos los lunes a las 7:00 AM** automáticamente.

### Opción B: Manual (cuando tú quieras)
1. Sube el nuevo Excel a la carpeta de SharePoint (reemplaza el anterior)
2. Ve a GitHub → **Actions → Actualizar Dashboard HGZ 1 → Run workflow**
3. En 2 minutos el dashboard está actualizado

---

## Estructura del Excel en SharePoint

El Excel debe tener exactamente el mismo formato que el reporte de IMSS:
- El encabezado real está en la **fila 5** (la fila 1 tiene el título, la 2 el periodo)
- Columnas de Microorganismo 1, 2, 3 y 4 con sus antibióticos
- El script maneja automáticamente múltiples periodos en el mismo archivo

### ¿Puedo juntar varios meses en un solo Excel?
**Sí.** El script detecta el periodo de cada fila por la fecha de detección de la infección. Puedes tener enero, febrero y marzo en un solo archivo y los separará automáticamente.

---

## Solución de problemas

| Problema | Solución |
|---|---|
| El workflow falla con error 401 | Los secretos de Azure están mal. Verifica SP_CLIENT_ID y SP_CLIENT_SECRET |
| El workflow falla con error 404 | El SP_FILE_PATH está mal. Verifica la ruta exacta del Excel en SharePoint |
| El dashboard no muestra datos | Verifica que el Excel tenga el formato correcto (encabezado en fila 5) |
| La URL de GitHub Pages no carga | Espera 5-10 minutos después de activar Pages por primera vez |
| Hay infecciones sin clasificar | El script las pone en "RESTO DE IAAS" y avisa en el log. Notifica al desarrollador para agregar el mapeo |

---

## Contacto para soporte técnico

Si algo falla, toma una captura del error en la pestaña **Actions** de GitHub y compártela.
