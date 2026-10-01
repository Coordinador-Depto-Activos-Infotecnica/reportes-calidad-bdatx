# Reportes BDATx

Proyecto de **calidad de datos** para la base **BDATx** (Coordinador Eléctrico Nacional). Genera reportes de relacionamiento, consistencia y escritura de las instalaciones, agrupados por tabla y por empresa; produce informes Word/PDF por empresa, un informe Word/PDF consolidado de toda la BDATx y reportes de **sugerencias de cambio** accionables (correcciones de nombre y posibles relacionamientos).

> **Importante:** este proyecto está pensado para ejecutarse en **Windows**, porque usa Microsoft Word para convertir los informes `.docx` a `.pdf`.

Si nunca has programado, no te preocupes: sigue la sección **Instalación paso a paso** en orden y copia los comandos tal cual.

---

## Índice

1. [Requisitos previos](#requisitos-previos)
2. [Instalación paso a paso](#instalación-paso-a-paso)
3. [Configuración del archivo `.env`](#configuración-del-archivo-env)
4. [Ejecución](#ejecución)
5. [Estructura del proyecto](#estructura-del-proyecto)
6. [Descripción de los scripts](#descripción-de-los-scripts)
7. [Salidas que genera](#salidas-que-genera)
8. [Solución de problemas](#solución-de-problemas)
9. [Notas técnicas](#notas-técnicas)
10. [Repositorio y contacto](#repositorio-y-contacto)

---

## Requisitos previos

Antes de empezar necesitas tener instalado lo siguiente:

| Programa | Para qué sirve | ¿Obligatorio? |
|----------|----------------|---------------|
| **Python 3.13 o superior** | Es el lenguaje en el que están escritos los programas. | Sí |
| **PostgreSQL 17** | Es la base de datos donde se cargan y consultan los datos de BDATx. Incluye la herramienta `psql.exe`. | Sí |
| **Microsoft Word** | Solo se usa para convertir los informes `.docx` a `.pdf`. | Sí, para generar los PDF |
| **Git** | Solo si vas a descargar el proyecto clonando el repositorio. | No (puedes usar el ZIP) |

Si ya tienes todo esto instalado, salta directo a [Configuración del archivo `.env`](#configuración-del-archivo-env).

---

## Instalación paso a paso

### Paso 1. Instalar Python 3.13

1. Entra a <https://www.python.org/downloads/>.
2. Descarga el **Windows installer (64-bit)** de la versión **3.13** o superior.
3. Ejecuta el instalador. En la primera pantalla, **marca la casilla “Add python.exe to PATH”** que aparece en la parte inferior. Es muy importante; si no la marcas, luego el comando `python` no funcionará.
4. Pulsa **Install Now** y espera a que termine.

### Paso 2. Comprobar que Python quedó instalado

1. Abre **PowerShell**: pulsa `Win + X` y elige **Windows PowerShell** o **Terminal**.
2. Escribe el siguiente comando y pulsa Enter:

   ```powershell
   python --version
   ```

   Debe mostrar algo parecido a `Python 3.13.7`.
3. Comprueba también `pip` (el instalador de librerías):

   ```powershell
   python -m pip --version
   ```

   Debe mostrar la versión de `pip`.

Si alguno de los dos comandos da error, revisa [Solución de problemas](#solución-de-problemas).

### Paso 3. Instalar PostgreSQL 17

1. Entra a <https://www.postgresql.org/download/windows/> y pulsa **Download the installer** (instalador de EDB).
2. Descarga **PostgreSQL 17** para Windows x86-64.
3. Ejecuta el instalador y deja las opciones por defecto. Asegúrate de que quede marcado **Command Line Tools**, porque allí viene `psql.exe`.
4. Cuando te pida una contraseña para el usuario **`postgres`**, escríbela y **guárdala**. Esa será tu `DB_PASSWORD`.
5. Deja el puerto por defecto **5432**.

### Paso 4. Comprobar que `psql` funciona

En PowerShell, ejecuta (respeta las comillas):

```powershell
& "C:\Program Files\PostgreSQL\17\bin\psql.exe" --version
```

Debe mostrar algo como `psql (PostgreSQL) 17.x`. Si el instalador usó otra ruta, anótala: la necesitarás en el archivo `.env` (variable `PSQL_PATH`).

> El servicio de PostgreSQL debe estar **iniciado** cada vez que uses el programa. Normalmente arranca solo con Windows. Si no, búscalo en **Servicios** (`Win + R` → `services.msc`) como `postgresql-x64-17` y pulsa **Iniciar**.

### Paso 5. Descargar el proyecto

Elige **una** de las dos opciones:

- **Con Git** (si lo tienes instalado):

  ```powershell
  git clone https://github.com/Coordinador-Depto-Activos-Infotecnica/reportes-calidad-bdatx.git
  cd reportes-calidad-bdatx
  ```

- **Sin Git**: abre la página del repositorio <https://github.com/Coordinador-Depto-Activos-Infotecnica/reportes-calidad-bdatx>, pulsa el botón verde **Code → Download ZIP**, descomprime el archivo y abre una terminal dentro de la carpeta del proyecto.

> Todos los comandos de aquí en adelante se ejecutan **dentro de la carpeta del proyecto**.

### Paso 6. Crear un entorno virtual (recomendado)

Un entorno virtual es una “caja” donde se instalan las librerías del proyecto sin afectar el resto de tu computador. Es opcional, pero muy recomendado.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

Al activarse verás `(.venv)` al inicio de la línea de la terminal.

Si PowerShell muestra un error como *“no se puede cargar el archivo Activate.ps1 porque la ejecución de scripts está deshabilitada”*, ejecuta esto una sola vez y vuelve a intentar:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

> Si prefieres el **Símbolo del sistema (CMD)** en lugar de PowerShell, activa el entorno con `.venv\Scripts\activate.bat`.

### Paso 7. Instalar las dependencias

Todas las librerías necesarias están listadas en el archivo **`requirements.txt`**. Instálalas con:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Esto descargará e instalará automáticamente:

- `psycopg2-binary` — conexión a PostgreSQL.
- `pandas`, `openpyxl`, `XlsxWriter` — manejo de datos y archivos Excel.
- `requests` — descarga de la base desde la API.
- `python-docx` — generación de los informes Word.
- `pywin32` — conversión de Word a PDF.

Si necesitas reinstalar todo desde cero más adelante, vuelve a ejecutar `pip install -r requirements.txt`.

### Paso 8. Crear y configurar el archivo `.env`

1. Crea una copia de la plantilla `.env_copy` con el nombre `.env`:

   ```powershell
   Copy-Item .env_copy .env
   ```

2. Ábrela con el Bloc de notas:

   ```powershell
   notepad .env
   ```

3. Completa los valores reales (ver la sección siguiente) y guarda el archivo.

### Paso 9. Primera ejecución

```powershell
python master_bdatx.py
```

El programa te hará una pregunta sobre el **origen de la base BDATx**:

- `n` → **no cargar** (usa la base que ya está en PostgreSQL).
- `d` → **descargar** la última base desde la API e instalarla.
- `a` → cargar una base guardada anteriormente en `BASE_FOLDER`.

La primera vez elige **`d`** para descargar los datos. El proceso puede tardar varios minutos.

---

## Configuración del archivo `.env`

El archivo `.env` guarda tus credenciales y rutas. **No se sube al repositorio** (está excluido en `.gitignore`), así que cada persona debe crear el suyo.

| Variable | Qué significa | Valor de ejemplo |
|----------|---------------|------------------|
| `DB_NAME` | Nombre de la base local donde se cargará BDATx. | `bdatx` |
| `DB_USER` | Usuario de PostgreSQL. | `postgres` |
| `DB_PASSWORD` | Contraseña del usuario anterior. | `MiClaveSecreta` |
| `DB_HOST` | Servidor donde está PostgreSQL. Normalmente la misma PC. | `localhost` |
| `DB_PORT` | Puerto de PostgreSQL. | `5432` |
| `PSQL_PATH` | Ruta completa al ejecutable `psql.exe`. | `C:\Program Files\PostgreSQL\17\bin\psql.exe` |
| `DOWNLOAD_URL` | Dirección desde donde se descarga la base. Ya viene configurada. | `https://...` |
| `BASE_FOLDER` | Carpeta donde se guardan las bases descargadas. | `C:/Users/Public/bdatx` |

Contenido típico del archivo `.env`:

```
# Base BDATx (PostgreSQL local)
DB_NAME=bdatx
DB_USER=postgres
DB_PASSWORD=MiClaveSecreta
DB_HOST=localhost
DB_PORT=5432

# Carga de la base
PSQL_PATH=C:\Program Files\PostgreSQL\17\bin\psql.exe
DOWNLOAD_URL=https://activos-tx-prod.appspot.com/api/v1/public-database-download/
BASE_FOLDER=C:/Users/Public/bdatx
```

Notas importantes:

- No es necesario crear la base `bdatx` a mano: el programa la **borra y la vuelve a crear** cada vez que carga una base nueva.
- Si tu contraseña contiene espacios o caracteres raros, escríbela **entre comillas** (por ejemplo `DB_PASSWORD="Mi clave 2026"`).
- En `BASE_FOLDER` puedes usar `/` (barra normal) o `\\` (doble barra invertida). Evita una sola barra invertida en Windows para esta variable.

---

## Ejecución

### Flujo principal

```powershell
python master_bdatx.py
```

Al ejecutarlo, el programa ofrece un menú:

```
¿Qué deseas ejecutar?
  1 = Flujo completo (todos los criterios)                      [por defecto]
  2 = Uno o más criterios + actualizar reportes
  3 = Solo actualizar reportes (Excel/Word/PDF)
```

- **Modo 1 (completo)**: pregunta el **origen de la base BDATx** (`n` no cargar / `d` descargar la última desde la API / `a` cargar una base guardada) y ejecuta todas las revisiones y luego los reportes.
- **Modo 2 (criterios)**: permite elegir revisiones numeradas (p. ej. `11`, `3,7`, `9-11` o `todos`) y luego actualiza los reportes. No pregunta el origen de la base (usa la base actual).
- **Modo 3 (solo salidas)**: ejecuta únicamente los reportes (Excel/Word/PDF) usando los resultados vigentes en `docs/reporte/`.

También admite ejecución no interactiva:

```powershell
python master_bdatx.py --listar        # lista las revisiones numeradas
python master_bdatx.py --revision 9-11 # uno o varios criterios + reportes
python master_bdatx.py --solo-salidas  # solo actualiza reportes
```

El flujo ejecuta, en orden:

1. `carga_bdatx.py` — (opcional) descarga la BD desde la API **o** carga una base guardada en `BASE_FOLDER`. Fija la fecha del snapshot en el comentario de la base (la usan los reportes).
2. `analisis_estse_oc_serv.py` — enriquecimiento de `oocc`, `estse` y `serv`.
3. `analisis_duplicados_desconexiones.py` — reportes de desconectados, duplicados y conexión nula.
4. `auditoria_lineas_metricas.py` — auditoría de la jerarquía de líneas.
5. `auditoria_contenedores_metricas.py` — contenedores vacíos.
6. `analisis_tension_patios.py` — coherencia de tensión en patios.
7. `revision_calidad_escriturageneral.py` — calidad de escritura de nombres.
8. `revision_herencia_nombre.py` — herencia del nombre de la subestación.
9. `analisis_prefijo.py` — verificación de prefijos de nombre.
10. `revision_abreviatura_torre.py` — abreviatura de línea en torres.
11. `revision_abreviatura_vano.py` — abreviatura de línea en vanos.
12. `revision_documentos_datasheet.py` — documentos obligatorios del datasheet de subestaciones y líneas.
13. `proceso_excel.py` — calcula `porcentaje error` y actualiza el avance diario.
14. `reporte_observaciones_empresa.py` — consolida las observaciones en Excel por empresa.
15. `sugerencias_cambios.py` — sugerencias de cambio (simples y complejas).
16. `formato_informe.py` — informes Word y PDF por empresa.
17. `informe_global.py` — informe Word/PDF consolidado de toda la BDATx y Excel resumen por empresa.

### Informes Word/PDF por empresa (independiente)

```powershell
python scripts/formato_informe.py
```

Genera un `.docx` y un `.pdf` por empresa en `docs/ArchivoReporteEmpresa/<AAAAMMDD_HHMMSS>_bd<AAAAMMDD>/{word,pdf}/`. Las métricas cuentan **instalaciones observadas** (tipo de instalación + `id`), no hallazgos: una instalación con varias observaciones cuenta una sola vez.

### Informe global BDATx (independiente)

```powershell
python scripts/informe_global.py
```

Genera, con las mismas métricas y secciones:

- Un único `.docx` y `.pdf` consolidado de todas las empresas (`Alcance: BDATx`).
- Un Excel `resumen_empresas.xlsx` con **una fila por empresa** (solo las que tienen hallazgos) y **una columna por métrica** (`Instalaciones observadas` y `% Error` de cada regla), más una fila `TOTAL BDATx`.

Salida en `docs/ArchivoReporteGlobal/<AAAAMMDD_HHMMSS>_bd<AAAAMMDD>/{word,pdf}/BDATx_global.{docx,pdf}` y `.../resumen_empresas.xlsx`.

### Generar el archivo de abreviaturas de línea

Los análisis de abreviatura de torre y vano leen el archivo `utils/abreviatura_linea.json`. Si no existe (o quieres actualizarlo), genéralo así:

```powershell
python utils/scripts/lineas_abreviadas.py
```

### Convertir un Markdown a Word (independiente)

```powershell
python utils/scripts/md2word.py "documento.md" --pdf
```

Convierte un `.md`/`.txt` a `.docx` (y, con `--pdf`, también a PDF) usando Microsoft Word. Si no se indica archivo, lo pide de forma interactiva.

### Ejecutar un script suelto

Cualquier análisis puede ejecutarse por separado, por ejemplo:

```powershell
python scripts/analisis_prefijo.py
```

---

## Estructura del proyecto

```
ReportesBDATx/
├── master_bdatx.py            # Orquestador del flujo principal
├── config.py                  # Carga .env y expone configuración y rutas
├── requirements.txt           # Lista de librerías a instalar
├── .env                       # Credenciales y parámetros (no versionado)
├── .env_copy                  # Plantilla de ejemplo (copiar como .env)
├── .gitignore
├── scripts/
│   ├── carga_bdatx.py                        # Descarga y carga la BD desde la API
│   ├── analisis_estse_oc_serv.py             # Enriquecimiento oocc / estse / serv
│   ├── analisis_duplicados_desconexiones.py  # Desconectados, duplicados, conexión nula
│   ├── analisis_prefijo.py                   # Verificación de prefijos de nombre
│   ├── analisis_tension_patios.py            # Coherencia de tensión patio/barra/paño
│   ├── auditoria_contenedores_metricas.py    # Contenedores vacíos
│   ├── auditoria_lineas_metricas.py          # Auditoría de jerarquía de líneas
│   ├── revision_calidad_escriturageneral.py  # Calidad de escritura
│   ├── revision_herencia_nombre.py           # Herencia del nombre de la subestación
│   ├── revision_abreviatura_torre.py         # Abreviatura de línea en torres
│   ├── revision_abreviatura_vano.py          # Abreviatura de línea en vanos
│   ├── revision_documentos_datasheet.py      # Documentos obligatorios del datasheet (anexos)
│   ├── proceso_excel.py                      # % error + avance diario
│   ├── reporte_observaciones_empresa.py      # Consolida observaciones en Excel por empresa
│   ├── sugerencias_cambios.py                # Sugerencias de cambio (simples y complejas)
│   ├── formato_informe.py                    # Informe Word y PDF por empresa
│   └── informe_global.py                     # Informe Word/PDF global + resumen Excel por empresa
├── utils/
│   ├── prefijos.json                           # Prefijos por tipo de instalación
│   ├── niveles_de_tension.json                 # Niveles de tensión y su letra
│   ├── nomenclatura_tipos_pano.json            # Letras de tipo de paño
│   ├── abreviaturas_ssee.json                  # Abreviaturas de nombres de subestación
│   ├── abreviatura_linea.json                  # Abreviatura oficial por línea (generado)
│   ├── relacionamiento_columnas_instalaciones.json  # Configuración de relacionamiento
│   ├── esqueletos_instalaciones_json/          # Esqueletos por tipo (campos del datasheet)
│   └── scripts/
│       ├── lineas_abreviadas.py                # Genera abreviatura_linea.json
│       └── md2word.py                          # Convierte un .md/.txt a Word (y PDF)
└── docs/                      # Salidas (no versionadas)
    ├── reporte/               # Reportes sin fecha
    ├── historico/             # Reportes con fecha
    ├── avance/                # Avance diario (avance_error_diario.xlsx)
    ├── ReporteExcelEmpresa/   # Observaciones consolidadas por empresa (Excel)
    ├── SugerenciasCambios/    # Sugerencias de cambio por corrida
    ├── ArchivoReporteEmpresa/ # Informes Word/PDF por empresa (una carpeta por corrida)
    └── ArchivoReporteGlobal/  # Informe Word/PDF global + resumen Excel por empresa
```

---

## Descripción de los scripts

| Script | Función | Salida principal |
|--------|---------|------------------|
| `carga_bdatx.py` | Descarga `database.psql.gz` desde la API **o** carga una base guardada en `BASE_FOLDER` (menú de bases disponibles), y la carga en la BD `bdatx` con `psql`. | BD `bdatx` |
| `analisis_estse_oc_serv.py` | Para `oocc`/`estse`/`serv` construye las tablas de relación y exporta archivos enriquecidos con hojas de relación, no relacionados, relacionados, maestra y resumen por empresa. | `api_{oocc,est_se,serv}_relacion_enriquecida.xlsx` |
| `analisis_duplicados_desconexiones.py` | Detecta registros **desconectados**, **duplicados** y con **conexión nula** por tabla relacionada. | `analisis_desconectados.xlsx`, `analisis_duplicados.xlsx`, `analisis_conexion_nula.xlsx` |
| `auditoria_lineas_metricas.py` | Construye el árbol Línea→Circuito→Tramo→Vano→Accesorio y aplica 10 reglas (contenedores vacíos, sufijos, herencia, accesorios, nodos, conectividad). | `auditoria_lineas_metricas.xlsx` |
| `auditoria_contenedores_metricas.py` | Detecta contenedores vacíos (subestación/patio por FK; paño/casa/armario por nodo compartido **o** por FK directa, p. ej. banco de baterías/generador hacia Casa SSGG) y torres/marcolineas sin OOCC ni accesorios. | `auditoria_contenedores_metricas.xlsx` |
| `analisis_tension_patios.py` | Verifica la coherencia del nivel de tensión entre patio, barra y paño. | `analisis_tension_patios.xlsx` |
| `revision_calidad_escriturageneral.py` | Dobles espacios, espacios extremos, `kV` mal escrito, minúsculas, guiones y separador ` - `. | `revision_calidad_escriturageneral.xlsx` |
| `revision_herencia_nombre.py` | Verifica que los hijos contengan el nombre de la subestación raíz. | `revision_herencia_nombre.xlsx` |
| `analisis_prefijo.py` | Verifica que el `name` comience con un prefijo válido definido en `utils/prefijos.json`. | `analisis_prefijo.xlsx` |
| `revision_abreviatura_torre.py` | Verifica que la abreviatura de línea del nombre de la torre coincida con `utils/abreviatura_linea.json`. | `revision_abreviatura_torre.xlsx` |
| `revision_abreviatura_vano.py` | Igual que torres, pero para los vanos. | `revision_abreviatura_vano.xlsx` |
| `revision_documentos_datasheet.py` | Revisa los campos de documento obligatorios del `datasheet` de subestaciones y líneas (`DOCUMENTO_FALTANTE`, `DOCUMENTO_NO_EXISTE`, `DOCUMENTO_INCORRECTO`). | `revision_documentos_datasheet.xlsx` |
| `proceso_excel.py` | Calcula `porcentaje error` en las hojas `analisis` y alimenta el avance diario (una columna por día). | `docs/avance/avance_error_diario.xlsx` |
| `reporte_observaciones_empresa.py` | Unifica las observaciones de todos los reportes en un Excel por empresa (con la columna `decreto`). | `docs/ReporteExcelEmpresa/<ts>/<empresa>.xlsx` |
| `sugerencias_cambios.py` | Sugerencias accionables: **simples** (corrección de nombre con columna `corregido`) y **complejas** (posible relacionamiento de OOCC/EstSE/Serv no relacionados). | `docs/SugerenciasCambios/<ts>/sugerenciascambios.xlsx` + uno por empresa |
| `formato_informe.py` | Consolida las métricas por empresa (instalaciones observadas) y genera un informe Word y PDF estilizado (banner + KPI + tres secciones). | `docs/ArchivoReporteEmpresa/<ts>/{word,pdf}/` |
| `informe_global.py` | Informe Word/PDF consolidado de toda la BDATx y Excel resumen con una fila por empresa (mismas métricas, en columnas). | `docs/ArchivoReporteGlobal/<ts>/{word,pdf}/BDATx_global.{docx,pdf}` + `resumen_empresas.xlsx` |
| `lineas_abreviadas.py` (`utils/scripts/`) | Genera `utils/abreviatura_linea.json` a partir de los nombres de las líneas de la BD. | `utils/abreviatura_linea.json` |
| `md2word.py` (`utils/scripts/`) | Convierte un archivo `.md`/`.txt` a Word (`.docx`) y, con `--pdf`, a PDF. | `generados/<archivo>.docx` |

---

## Salidas que genera

- `docs/reporte/` — reportes vigentes (sin fecha).
- `docs/historico/` — copias con fecha/hora (`YYYY-MM-DD_HH-MM-SS`) y la fecha de la base BDATx (`-bdYYYY-MM-DD`).
- `docs/avance/avance_error_diario.xlsx` — seguimiento diario de la mejora (columna por fecha de la base BDATx).
- `docs/ReporteExcelEmpresa/<ts>_bd<AAAAMMDD>/` — observaciones consolidadas en Excel por empresa.
- `docs/SugerenciasCambios/<ts>_bd<AAAAMMDD>/` — sugerencias de cambio (consolidado y por empresa).
- `docs/ArchivoReporteEmpresa/<ts>_bd<AAAAMMDD>/` — informes Word y PDF por empresa.
- `docs/ArchivoReporteGlobal/<ts>_bd<AAAAMMDD>/` — informe Word/PDF global y `resumen_empresas.xlsx` (una fila por empresa).

Toda la carpeta `docs/` se crea automáticamente y está excluida del control de versiones.

---

## Solución de problemas

**`python : no se reconoce como nombre de un comando...`**
Python no está en el `PATH`. Vuelve a ejecutar el instalador de Python, elige **Modify** y marca *“Add python.exe to PATH”*. Alternativamente prueba con el lanzador `py --version` o reinstala marcando la casilla del Paso 1.

**`pip : no se reconoce como nombre de un comando...`**
Usa siempre la forma larga: `python -m pip install -r requirements.txt`.

**`.venv\Scripts\Activate.ps1 ... la ejecución de scripts está deshabilitada`**
Ejecuta una vez `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` y acepta con `S`. O usa CMD con `.venv\Scripts\activate.bat`.

**`ModuleNotFoundError: No module named 'psycopg2'` (o `pandas`, `docx`, `win32com`)**
El entorno virtual no está activado o las dependencias no se instalaron. Activa el entorno (`(.venv)` debe verse al inicio de la línea) y ejecuta de nuevo `pip install -r requirements.txt`.

**`psql.exe` no se reconoce / “El sistema no puede encontrar la ruta especificada”**
La ruta en `PSQL_PATH` no es correcta. Compruébala en PowerShell:

```powershell
Test-Path "C:\Program Files\PostgreSQL\17\bin\psql.exe"
```

Debe responder `True`. Si responde `False`, busca dónde quedó instalado `psql.exe` y corrige `PSQL_PATH`.

**`password authentication failed for user "postgres"`**
La contraseña en `DB_PASSWORD` es incorrecta. Corrígela en `.env`.

**`could not connect to server: Connection refused` / `Is the server running...`**
El servicio de PostgreSQL no está iniciado. Ábrelo con `Win + R` → `services.msc`, busca `postgresql-x64-17` e inícialo.

**La revisión de abreviaturas aparece vacía o con `[WARN] No existe ...abreviatura_linea.json`**
Genera el archivo con `python utils/scripts/lineas_abreviadas.py`.

**Error al generar PDF (`Dispatch`, `Word.Application` o `pywin32`)**
Instala `pywin32` (`pip install -r requirements.txt`) y verifica que Microsoft Word esté instalado y abra correctamente.

**Aparece una advertencia de certificado SSL (`InsecureRequestWarning`)**
Es normal durante la descarga de la base; no impide el proceso.

---

## Notas técnicas

- `docs/`, `.env`, `generados/`, `__pycache__/` y `*.pyc` están excluidos del control de versiones (`.gitignore`).
- Los análisis de calidad se calculan sobre registros con `status = 'EN_OPERACION'`.
- El nombre de empresa se obtiene de `propietario_id` → `api_empresa.id` → `api_empresa.name`.
- La fecha de la base BDATx se guarda en el `COMMENT ON DATABASE bdatx` (ej. `BDATx 2026-09-07 16:14:00`) y se expone con `config.fecha_bd()`; los reportes la agregan como sufijo `-bd<fecha>`.
- La conversión de Word a PDF usa `win32com.client` y requiere Microsoft Word instalado.
- El archivo `utils/abreviatura_linea.json` se regenera con `python utils/scripts/lineas_abreviadas.py`; si no existe, las revisiones de abreviatura de torre y vano se omiten con una advertencia.
- `utils/esqueletos_instalaciones_json/` contiene los esqueletos por tipo de instalación; `revision_documentos_datasheet.py` toma de ahí los campos de documento obligatorios del `datasheet` de subestaciones y líneas.

---

## Repositorio y contacto

- **Repositorio público**: <https://github.com/Coordinador-Depto-Activos-Infotecnica/reportes-calidad-bdatx>
- **Consultas**: cualquier consulta o requerimiento debe dirigirse a **activostx@coordinador.cl**.
