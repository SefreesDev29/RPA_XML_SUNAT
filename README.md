# 📑 Robot Descarga XMLs de SUNAT

> **Descripción:** Consulta automatizada de comprobantes electrónicos en el portal SOL de SUNAT y descarga masiva de archivos XML.

Herramienta diseñada para el procesamiento en lote de comprobantes (Facturas y Notas de Crédito), con ejecución paralela por área (`Siniestros` y `Salud`), tolerancia a fallos, recuperación automática de sesión, descompresión de paquetes ZIP, envío de correos vía Microsoft Graph API y auditoría de ejecución.

---

## 🚀 Tecnologías y Requisitos

![Python](https://img.shields.io/badge/Python-3.12%2B-3776AB.svg?style=for-the-badge&logo=python&logoColor=white)
![uv](https://img.shields.io/badge/Gestor_de_Paquetes-uv-D042ED.svg?style=for-the-badge)

> **Compatibilidad:** Requiere **Python `>=3.11`** (desarrollado y validado en **Python `3.14`** sobre **Windows**).

### 📚 Stack Principal

**Automatización y Web Scraping:**  
![Playwright](https://img.shields.io/badge/playwright-1.64.1-2EAD33.svg?style=flat-square&logo=playwright&logoColor=white)

**Procesamiento de Datos:**  
![Pandas](https://img.shields.io/badge/pandas-%3E%3D3.1.0-054593.svg?style=flat-square&logo=pandas&logoColor=white)

**Manejo de Archivos Excel:**  
![Openpyxl](https://img.shields.io/badge/openpyxl-%3E%3D3.1.5-217346.svg?style=flat-square&logo=openpyxl)
![XlsxWriter](https://img.shields.io/badge/xlsxwriter-%3E%3D3.3.0-217346.svg?style=flat-square&logo=xlsxwriter)

**Envío de Correo:**  
![MSAL](https://img.shields.io/badge/msal-%3E%3D1.39.0-CA3519.svg?style=flat-square&logo=msal)
![Requests](https://img.shields.io/badge/requests-%3E%3D2.34.2-2CA5E0.svg?style=flat-square)

**Configuración y Credenciales:**  
![Keyring](https://img.shields.io/badge/keyring-%3E%3D25.7.0-015F78.svg?style=flat-square&logo=keyring)
![Pydantic-Settings](https://img.shields.io/badge/pydantic--settings-%3E%3D2.16.1-4D179A.svg?style=flat-square&logo=pydantic--settings)

**Consola y Trazabilidad:**  
![Loguru](https://img.shields.io/badge/loguru-%3E%3D0.7.3-43B02A.svg?style=flat-square&logo=loguru)
![Rich](https://img.shields.io/badge/rich-%3E%3D15.0.0-99960F.svg?style=flat-square&logo=rich&logoColor=white)

---

## 🛠️ Guía de Instalación y Uso

### 1️⃣ Instalar `pipx` y `uv`
Para gestionar las dependencias de manera rápida y global, instala `uv` utilizando `pipx`:

```bash
# 1. Instalar pipx
python -m pip install --user pipx --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org

# 2. Agregar pipx a la variable PATH del sistema
python -m pipx ensurepath

# 3. Instalar uv
pipx install uv --pip-args="--trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org"
```

### 2️⃣ Configurar el Entorno del Proyecto
> ⚠️ **Nota Importante:** Este paso **solo se ejecuta una vez** (cuando clonas el repositorio o instalas el proyecto desde cero).

Una vez con `uv` instalado, prepárate para ejecutar el proyecto creando su entorno aislado:

```bash
# 1. Crear el entorno virtual
uv venv

# 2. Inicializar configuración (si se requiere desde cero)
uv init

# 3. Sincronizar dependencias (instala todo lo de pyproject.toml)
uv sync

# (Opcional) Verificar librerías instaladas
uv pip list
```

### 3️⃣ Configurar Variables de Entorno (`.env`) y Credenciales de Windows
Copiar [`.env.example`](.env.example) como `.env` en la raíz del proyecto y completar los valores reales:

```bash
# Colocar los *valores reales* de cada variable para un corecto funcionamiento
SUNAT_RUC_LPG=20XXXXXXXXX
SUNAT_RUC_LPE=20XXXXXXXXX
PATH_CLOUD_TEST=C:\Users\user_qa\...
PATH_CLOUD=C:\Users\user_prod\...
PATH_CLOUD_TEST=C:\Users\usuario\OneDrive - Empresa\Carpeta_Pruebas
PATH_CLOUD=C:\Users\usuario\OneDrive - Empresa\Carpeta_Produccion
MAIL_BOX=desarrollo@empresa.com.pe
RECIPIENTS_MAIL_LIST=usuario@empresa.com.pe|operario@empresa.com.pe
```

> ⚠️ **Aprovisionamiento de secretos (`supply_manager.bat`):** Ejecutar **`supply_manager.bat`** para registrar en **Windows Credential Manager** (`RPA_SUNAT_DEV`, `RPA_SUNAT_TEST` o `RPA_SUNAT_PROD`) los **7 secretos** obligatorios: `USER_LPG`, `PASS_LPG`, `USER_LPE`, `PASS_LPE`, `TENANT_ID`, `CLIENT_ID` y `CLIENT_SECRET`.

### 4️⃣ Ejecución del Robot

Colocar el archivo Excel con los comprobantes a consultar en la carpeta de entrada de cada área en OneDrive:

```bash
# Ruta de Carpeta Sincronizada
f"{PATH_CLOUD}/Inputs/"

# Ruta de entrada en OneDrive por área (Siniestros | Salud)
{PATH_CLOUD}\{Area}\XMLs_SUNAT\Inputs\

# Columnas requeridas
(EMPRESA, RUC, COD_DCMNTO, SERIEDOC, NUMERODOC)
```

> ⚠️ **Nota Importante:** La columna `COD_DCMNTO` acepta valores numéricos o texto de 2 dígitos (`01`, `03`, `07`, `08`, `R1`).

Para iniciar el robot, ejecutar siguiente comando:
#### Modos de Ejecución
* **Lanzador Producción (`APP_ENV=PROD`, ambas áreas en paralelo):** ejecutar `run_robot.bat`.
* **Lanzador Pruebas (`APP_ENV=TEST`, ambas áreas en paralelo):** ejecutar `run_test.bat`.
* **Aprovisionador de Credenciales:** ejecutar `supply_manager.bat`.
* **Ejecución Manual por CLI (`APP_ENV=DEV` por defecto):**
  ```powershell
  # Ejecutar todas las áreas en paralelo
  uv run Main.py

  # Ejecutar un área específica
  uv run Main.py Salud
  uv run Main.py Siniestros
  ```

### 📂 Estructura Operativa Multi-Área

```
- Areas_Local/{Area}/Inputs/: Cola local activa (mantiene pendientes/remanentes para reintento).
- Areas_Local/{Area}/Outputs/Xmls/: Carpeta temporal de XMLs del ciclo (limpiada tras respaldo en OneDrive).
- Areas_Local/{Area}/Outputs/Procesados/: Staging local de reportes de descargados antes de copiar a OneDrive.
- Areas_Local/{Area}/Outputs/Observados/: Staging local de reportes observados antes de copiar a OneDrive.
- Areas_Local/{Area}/Perfil_Edge/: Perfil persistente de Microsoft Edge aislado por área.
- {PATH_CLOUD}/{Area}/XMLs_SUNAT/Historicos/: Respaldo de Excel ingeridos desde OneDrive.
- {PATH_CLOUD}/{Area}/XMLs_SUNAT/Xmls/: Lotes de XMLs descomprimidos (excluyendo constancias CDR R-).
- {PATH_CLOUD}/{Area}/XMLs_SUNAT/Comprimidos_Correo/: Consolidados ZIP enviados por correo.
- {PATH_CLOUD}/{Area}/XMLs_SUNAT/Procesados/: Reporte Excel formateado como tabla con descargados exitosos.
- {PATH_CLOUD}/{Area}/XMLs_SUNAT/Observados/: Reportes Excel con registros inválidos o no encontrados en SUNAT.
- Logs/Log_Robot_Sunat_{Area}_{YYYYMMDD}.log: Bitácoras diarias de auditoría por área.
```
