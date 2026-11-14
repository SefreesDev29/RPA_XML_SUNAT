# 📑 Robot Descarga XMLs de SUNAT

> **Descripción:** Consulta automatizada de comprobantes electrónicos en el portal SOL de SUNAT y descarga masiva de archivos XML.

Herramienta diseñada para el procesamiento en lote de comprobantes (Facturas y Notas de Crédito), con tolerancia a fallos, recuperación automática de sesión, descompresión de paquetes ZIP, envíos de correos y auditoría de ejecución.

---

## 🚀 Tecnologías y Requisitos

![Python](https://img.shields.io/badge/Python-3.14%2B-3776AB.svg?style=for-the-badge&logo=python&logoColor=white)
![uv](https://img.shields.io/badge/Gestor_de_Paquetes-uv-D042ED.svg?style=for-the-badge)

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

**Envío de Variables de Entorno y Credenciales:**  
![Keyring](https://img.shields.io/badge/keyring-%3E%3D25.7.0-015F78.svg?style=flat-square&logo=keyring)
![Pydantic-Settings](https://img.shields.io/badge/pydantic--settings-%3E%3D2.16.1-4D179A.svg?style=flat-square&logo=pydantic--settings)

**Consola y Trazabilidad:**  
![Loguru](https://img.shields.io/badge/loguru-%3E%3D0.7.3-43B02A.svg?style=flat-square&logo=loguru)
![Rich](https://img.shields.io/badge/rich-%3E%3D15.0.0-99960F.svg?style=flat-square&logo=rich&logoColor=white)

---

## 🛠️ Guía de Instalación y Uso

### 1️⃣ Instalar `pipx` y `uv`
Para gestionar las dependencias de manera rápida y global, instalaremos `uv` utilizando `pipx`. Ejecuta lo siguiente en tu terminal:

```bash
# 1. Instalar pipx
python -m pip install --user pipx --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org

# 2. Agregar pipx a la variable PATH del sistema
python -m pipx ensurepath

# 3. Instalar uv
pipx install uv --pip-args="--trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org"
```

### 2️⃣ Configurar el Entorno del Proyecto
> ⚠️ **Nota Importante:** Este paso **solo se ejecuta una vez** (cuando clonas el repositorio o instalas el proyecto desde cero). No es necesario repetirlo en cada ejecución.

Una vez con `uv` instalado, prepárate para ejecutar el proyecto creando su entorno aislado:

```bash
# 1. Crear el entorno virtual
uv venv

# 2. Inicializar configuración (si se requiere desde cero)
uv init

# 3. Sincronizar dependencias (instala todo lo de pyproject.toml)
uv sync --upgrade

# (Opcional) Verificar que todas las librerías se hayan instalado correctamente listándolas con:
uv pip list
```

### 3️⃣ Configurar Variables de Entorno (.env) y Credenciales de Windows
Crear un archivo .env en la raíz del proyecto con la siguiente estructura:

```bash
# Colocar los *valores reales* de cada variable para un corecto funcionamiento
SUNAT_RUC_LPG=20XXXXXXXXX
SUNAT_RUC_LPE=20XXXXXXXXX
PATH_CLOUD_TEST=C:\Users\user_qa\...
PATH_CLOUD=C:\Users\user_prod\...
```

> ⚠️ **Nota Importante:** Ejecutar el archivo **supply_manager.bat** para registrar las credenciales de SUNAT de manera segura en Windows.

### 4️⃣ Ejecución del Robot

Colocar el archivo Excel con los comprobantes a consultar en la carpeta:

```bash
# Ruta de Carpeta Sincronizada
f"{PATH_CLOUD}/Inputs/"

# Columnas requeridas
(EMPRESA, RUC, COD_DCMNTO, SERIEDOC, NUMERODOC)
```

> ⚠️ **Nota Importante:** La columna COD_DCMNTO es de **tipo numérico (entero)** por lo que se espera que tenga los códigos del tipo de comprobante (Ejm. 01 o 07).

Para iniciar el robot, ejecutar siguiente comando:

```bash
uv run Main.py
```

Estructura y Resultados generados:

```
- Inputs/: Mantiene únicamente los registros nuevos o pendientes de reintento si el portal experimenta intermitencias durante la ejecución.

- Outputs/Xmls/: Archivos .xml descomprimidos (excluyendo las constancias CDR R-).

- Outputs/Procesados/: Reporte Excel formateado como tabla con los comprobantes descargados exitosamente.

- Outputs/Observados/: Reporte Excel con los comprobantes rechazados por carecer de RUC válido, campos obligatorios o que no se encontró información en el portal SUNAT.

- Logs/: Archivos de auditoría en tiempo real.
```

