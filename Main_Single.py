import random
import re
import time
import zipfile
from rich.console import Console
from Config import settings
from loguru import logger
from typing import Callable
from playwright.sync_api import sync_playwright, Playwright, Page, Browser, BrowserContext, FrameLocator

from Utils import (
    Path,
    print,
    add_log_console,
    add_log_file,
    export_df_to_excel,
    get_current_datetime,
    copy_items_to_folder,
    control_storage_day,
    logger,
    datetime,
    pd,
    remove_log,
    show_custom_rule,
    start_log,
    sys, 
    shutil
)

START_TIME, END_TIME = get_current_datetime(), get_current_datetime()
PERIOD_TIME = f"{START_TIME.year}{START_TIME.month:02d}{START_TIME.day:02d}"
PATH_GENERAL = Path(__file__).resolve().parent
PATH_GENERAL_CLOUD = Path(settings.PATH_CLOUD_TEST)

PATH_FOLDER_HISTORY_CLOUD = PATH_GENERAL_CLOUD / "Historicos"
PATH_FOLDER_INPUTS_CLOUD = PATH_GENERAL_CLOUD / "Inputs"
PATH_FOLDER_XMLS_CLOUD = PATH_GENERAL_CLOUD / "Xmls" / f"Proceso_{PERIOD_TIME}"
PATH_FOLDER_XMLS_ZIP_CLOUD = PATH_GENERAL_CLOUD / "Comprimidos_Correo"
PATH_FOLDER_PROCESSED_CLOUD = PATH_GENERAL_CLOUD / "Procesados"
PATH_FOLDER_QUARANTINE_CLOUD = PATH_GENERAL_CLOUD / "Observados"

PATH_FOLDER_INPUTS = PATH_GENERAL / "Inputs"
PATH_FOLDER_PROCESSED = PATH_GENERAL / "Outputs" / "Procesados"
PATH_FOLDER_QUARANTINE = PATH_GENERAL / "Outputs" / "Observados"
PATH_FOLDER_XMLS = PATH_GENERAL / "Outputs" / "Xmls" / f"Proceso_{PERIOD_TIME}"
PATH_FOLDER_XMLS_ZIP = PATH_GENERAL / "Outputs" / "Xmls" / f"Comprimidos_{PERIOD_TIME}"

PATH_FOLDER_LOGS = PATH_GENERAL / "Logs"
PATH_LOGS = PATH_FOLDER_LOGS / f"Log_Robot_Sunat_{PERIOD_TIME}.log"
FILE_LOG_EXISTS = False
PATH_LOCK_FILE = PATH_GENERAL / "Robot.lock"

for folder in [
    PATH_FOLDER_INPUTS,
    PATH_FOLDER_LOGS,
    PATH_FOLDER_PROCESSED,
    PATH_FOLDER_QUARANTINE,
    PATH_FOLDER_XMLS,
]:
    folder.mkdir(parents=True, exist_ok=True)

SUNAT_RUC_LPG = settings.SUNAT_RUC_LPG
SUNAT_USER_PROD_LPG = settings.SUNAT_USER_PROD_LPG
SUNAT_PASSWORD_PROD_LPG = settings.SUNAT_PASSWORD_PROD_LPG

SUNAT_RUC_LPE = settings.SUNAT_RUC_LPE
SUNAT_USER_PROD_LPE = settings.SUNAT_USER_PROD_LPE
SUNAT_PASSWORD_PROD_LPE = settings.SUNAT_PASSWORD_PROD_LPE

PATH_PERFIL_EDGE = PATH_GENERAL / "Perfil_Edge_Robot"

URL_LOGIN_SUNAT = (
    "https://api-seguridad.sunat.gob.pe/v1/clientessol/4f3b88b3-d9d6-402a-b85d-6a0bc857746a/oauth2/loginMenuSol"
    "?lang=es-PE&showDni=true&showLanguages=false"
    "&originalUrl=https://e-menu.sunat.gob.pe/cl-ti-itmenu/AutenticaMenuInternet.htm"
    "&state=rO0ABXNyABFqYXZhLnV0aWwuSGFzaE1hcAUH2sHDFmDRAwACRgAKbG9hZEZhY3RvckkACXRocmVzaG9sZHhwP0AAAAAAAAx3CAAAABAAAAADdAADZXhlcHQABnBhcmFtc3QASyomKiYvY2wtdGktaXRtZW51L01lbnVJbnRlcm5ldC5odG0mYjY0ZDI2YThiNWFmMDkxOTIzYjIzYjY0MDdhMWMxZGI0MWU3MzNhNnQABGV4ZWNweA=="
)

MAX_RECORDS_PER_RUN = 1500
BREAK_MIN_CALL = 2.0
BREAK_MAX_CALL = 2.0
BATCH_SIZE = 500
BREAK_SECONDS_RECOVER = 3
BREAK_MINUTE_BATCH = 1.0
MAX_RETRIES_PER_RECORD = 3

MAP_TYPES_CPE = {
    "01": "Factura",
    "03": "Boleta de Venta",
    "07": "Factura - Nota de Crédito",
    "08": "Factura - Nota de Débito",
    "R1": "Recibo por Honorarios",
}

COLUMNS_FINAL = ["ARCHIVO_ORIGEN", "EMPRESA", "#EXP", "RUC", "COD_DCMNTO", "TPO_DCMNTO", "SERIEDOC", "NUMERODOC", "RAZON_SOCIAL", "ESTADO", "DETALLE_ESTADO", "FECHA_PROCESO"]

class SunatConexionError(Exception):
    pass

class Robot_Sunat:
    def __init__(self):
        self.Execute_Robot()

    def Acquire_Lock(self) -> bool:
        if PATH_LOCK_FILE.exists():
            tiempo_creacion = PATH_LOCK_FILE.stat().st_mtime
            if time.time() - tiempo_creacion > 12600:
                PATH_LOCK_FILE.unlink(missing_ok=True)
            else:
                return False

        PATH_LOCK_FILE.write_text(str(time.time()), encoding="utf-8")
        return True

    def Release_Lock(self) -> None:
        PATH_LOCK_FILE.unlink(missing_ok=True)

    def Random_Delay(self):
        time.sleep(random.uniform(BREAK_MIN_CALL, BREAK_MAX_CALL))

    def Index_XMLs_Locals(self) -> dict[tuple[str, str], str]:
        indice_xmls: dict[tuple[str, str], str] = {}
        if not PATH_FOLDER_XMLS.exists():
            return indice_xmls

        patron_cpe = re.compile(r"([A-Za-z0-9]{4})-(\d+)\.xml$", re.IGNORECASE)

        for archivo in PATH_FOLDER_XMLS.glob("*.xml"):
            nombre = archivo.name
            match = patron_cpe.search(nombre)
            if match:
                serie_archivo = self.Normalize_Serie(
                    match.group(1).upper().strip()
                )
                numero_raw = match.group(2).strip()

                if len(numero_raw) > 11:
                    numero_raw = numero_raw[:-11]

                numero_archivo = (
                    str(int(numero_raw))
                    if numero_raw.isdigit()
                    else numero_raw
                )
                indice_xmls[(serie_archivo, numero_archivo)] = nombre

        logger.info(
            f"Índice local cargado: {len(indice_xmls)} XMLs encontrados previamente en disco."
        )
        return indice_xmls

    def Normalize_Serie(self, serie_raw: str) -> str:
        s = serie_raw.strip().replace(" ", "").upper()
        serie_final = s
        if len(s) > 1 and s.startswith("0") and s[1].isalpha():
            serie_final = s[1:]
        elif len(s) > 1 and s.startswith("N"):
            serie_final = "F" + s[1:]
        return serie_final[:4] if len(serie_final) > 4 else serie_final

    def Load_Clean_Data(self) -> tuple[pd.DataFrame, pd.DataFrame, list[Path]]:
        if not PATH_FOLDER_INPUTS.exists():
            raise Exception(f"No se encontró carpeta a procesar.\nUbicación Carpeta: {PATH_FOLDER_INPUTS}")

        excel_files = list(PATH_FOLDER_INPUTS.glob("*.xlsx"))

        if not excel_files:
            raise Exception(f"No se encontraron archivos Excel para procesar.\nUbicación Carpeta Cloud: {PATH_FOLDER_INPUTS_CLOUD}\nUbicación Carpeta Local: {PATH_FOLDER_INPUTS}")

        dfs: list[pd.DataFrame] = []
        
        for file in excel_files:
            try:
                temp_df = pd.read_excel(file, dtype=str)

                temp_df.columns = [c.strip().upper() for c in temp_df.columns]

                if "ARCHIVO_ORIGEN" not in temp_df.columns:
                    temp_df["ARCHIVO_ORIGEN"] = file.name

                if "EMPRESA" not in temp_df.columns:
                    temp_df["EMPRESA"] = "LPG"

                for col in ["#EXP", "RUC", "COD_DCMNTO", "TPO_DCMNTO", "SERIEDOC", "NUMERODOC", "RAZON_SOCIAL"]:
                    if col not in temp_df.columns:
                        temp_df[col] = ""
                    temp_df[col] = (
                        temp_df[col].fillna("").astype(str).str.strip()
                    )

                dfs.append(temp_df)
                logger.info(f"Cargado archivo: {file.name} ({len(temp_df)} filas)")
            except Exception as e:
                logger.error(f"Error al leer archivo excel. Ubicación Excel: '{file.parent}/{file.name}'.\n{e}")

        if not dfs:
            raise Exception(f"No se encontraron archivos Excel para procesar.\nUbicación Carpeta: {PATH_FOLDER_INPUTS}")

        df_consolidado = pd.concat(dfs, ignore_index=True)

        df_consolidado["RUC"] = df_consolidado["RUC"].str.replace(" ", "").replace(",", "").str.split(".").str[0]
        df_consolidado["SERIEDOC"] = df_consolidado["SERIEDOC"].apply(self.Normalize_Serie)
        df_consolidado["COD_DCMNTO"] = df_consolidado["COD_DCMNTO"].apply(
            lambda x: str(x).zfill(2) if str(x).isdigit() and len(str(x)) == 1 else x
        )

        rows_preview = len(df_consolidado)
        df_consolidado = df_consolidado.drop_duplicates(
            subset=["RUC", "COD_DCMNTO", "SERIEDOC", "NUMERODOC"]
        ).copy()
        if len(df_consolidado) < rows_preview:
            logger.info(
                f"Se descartaron {rows_preview - len(df_consolidado)} registros duplicados."
            )

        is_ruc_valid = df_consolidado["RUC"].str.match(r"^\d{11}$")
        fields_complete = (
                (df_consolidado["SERIEDOC"] != "") & 
                (df_consolidado["SERIEDOC"].str.len() > 3) & 
                (df_consolidado["NUMERODOC"] != "")
        )
        mask_valid = is_ruc_valid & fields_complete

        df_validos = df_consolidado[mask_valid].copy()
        df_observados = df_consolidado[~mask_valid].copy()

        df_validos["ESTADO"] = "PENDIENTE"
        df_validos["DETALLE_ESTADO"] = ""
        df_validos["FECHA_PROCESO"] = pd.Timestamp.now()

        df_observados["ESTADO"] = "OBSERVADO"
        df_observados["DETALLE_ESTADO"] = df_observados.apply(
            lambda row: (
                "RUC_VACÍO_O_INVÁLIDO"
                if not re.match(r"^\d{11}$", str(row["RUC"]))
                else "SERIE_VACÍO_O_INVÁLIDO"
                if (row["SERIEDOC"] == "" or len(row["SERIEDOC"]) <= 3)
                else "NÚMERO_VACÍO_O_INVÁLIDO"
            ),
            axis=1
        )
        df_observados["FECHA_PROCESO"] = pd.Timestamp.now()

        return df_validos, df_observados, excel_files

    def Wait_Condition(self, page: Page, condition: Callable[[], bool], timeout_ms: int = 25000, interval_ms: int = 150) -> bool:
        limite = time.time() + (timeout_ms / 1000.0)
        while time.time() < limite:
            if condition():
                return True
            page.wait_for_timeout(interval_ms)
        return False

    def Login_Sunat(self, playwright_instance: Playwright, max_login_attempts: int = 3):
        for attempt in range(1, max_login_attempts + 1):
            context = None
            try:
                logger.info(f"Lanzando navegador Edge y autenticando en SOL (Intento {attempt}/{max_login_attempts})...")

                context = playwright_instance.chromium.launch_persistent_context(
                    user_data_dir=str(PATH_PERFIL_EDGE),
                    channel="msedge",
                    headless=False,
                    chromium_sandbox=True,
                    ignore_default_args=["--no-sandbox"],
                    no_viewport=True,
                    accept_downloads=True,
                    args=[
                        "--start-maximized",
                        "--disable-blink-features=AutomationControlled",
                        "--no-first-run",
                        "--no-default-browser-check",
                    ],
                )

                page = context.pages[0] if context.pages else context.new_page()

                page.goto(URL_LOGIN_SUNAT, wait_until="networkidle", timeout=30000)

                field_ruc = page.get_by_role("textbox", name="RUC")
                field_ruc.click()
                field_ruc.press_sequentially(SUNAT_RUC_LPG, delay=5)

                field_user = page.get_by_role("textbox", name="Usuario")
                field_user.click()
                field_user.press_sequentially(SUNAT_USER_PROD_LPG, delay=15)

                field_pass = page.get_by_role("textbox", name="Contraseña")
                field_pass.click()
                field_pass.press_sequentially(SUNAT_PASSWORD_PROD_LPG, delay=15)

                page.wait_for_timeout(300)

                page.get_by_role("button", name="Iniciar sesión").click()

                page.wait_for_url("**/MenuInternet.htm*", timeout=30000)
                logger.info("Credenciales aceptadas por SUNAT...")

                page.wait_for_timeout(300)

                tree_menu = [
                    ("#divOpcionServicio1", "Módulo Empresas"),
                    ("#nivel1_11", "Comprobantes de Pago"),
                    ("#nivel2_11_38", "SEE - Facturas y Recibos"),
                    ("#nivel3_11_38_1", "Factura Electrónica"),
                    (
                        "#nivel4_11_38_1_1_1",
                        "Nueva Consulta de Comprobantes",
                    ),
                ]

                for selector, _ in tree_menu:
                    nodo = page.locator(selector)
                    nodo.wait_for(state="visible", timeout=10000)
                    nodo.click()
                    page.wait_for_timeout(100)

                frame = page.frame_locator('iframe[name="iframeApplication"]')

                logger.info("Esperando carga del formulario de consulta dentro del iframe...")

                modal_form = (
                    frame.get_by_text("Recibido")
                    .or_(frame.locator("input[name='rucEmisor']"))
                    .first
                )
                modal_error_frame = (
                    frame.get_by_text("servicios de SUNAT")
                    .or_(frame.get_by_text("reintentar en 5 minutos"))
                    .first
                )
                modal_error_page = (
                    page.get_by_text("servicios de SUNAT")
                    .or_(page.get_by_text("reintentar en 5 minutos"))
                    .first
                )

                ready = self.Wait_Condition(
                    page=page,
                    condition=lambda: (
                        modal_form.is_visible()
                        or modal_error_frame.is_visible()
                        or modal_error_page.is_visible()
                    ),
                    timeout_ms=20000,
                    interval_ms=100
                )
    
                if not ready:
                    raise Exception("TIMEOUT_CARGA_IFRAME_FORMULARIO")

                if modal_error_frame.is_visible() or modal_error_page.is_visible():
                    logger.error(f"Detectado 'Error del Servidor' de SUNAT en intento {attempt}...")
                    raise Exception("SUNAT_CAIDO_AL_INICIAR: Error del Servidor detectado.")

                logger.info(
                    "Formulario de consulta cargado y listo para operar..."
                )
                return context, page, frame
            except Exception as e:
                logger.warning(
                    f"Fallo en intento de inicio de sesión {attempt}/{max_login_attempts}.\n{e}"
                )
                if context:
                    try:
                        context.close()
                    except Exception:
                        pass

                if attempt < max_login_attempts:
                    logger.info("Esperando unos segundos para que los servicios de SUNAT respondan...")
                    time.sleep(BREAK_SECONDS_RECOVER)
                else:
                    raise SunatConexionError(
                        f"SUNAT inaccesible tras {max_login_attempts} intentos."
                    )

    def Recover_Screen_Restart(self, playwright_instance: Playwright, context: BrowserContext, page: Page):
        logger.warning("Iniciando protocolo de recuperación de pantalla...")

        try:
            btn_new_search = page.get_by_text("Nueva Consulta de").first
            if btn_new_search.is_visible():
                btn_new_search.click(timeout=5000)
                page.wait_for_timeout(1500)

                frame = page.frame_locator('iframe[name="iframeApplication"]')
                frame.get_by_text("Recibido").wait_for(
                    state="visible", timeout=15000
                )
                logger.info("Recuperación Nivel 1 (Soft Reset) exitosa...")
                return context, page, frame
        except:
            logger.debug(f"Soft Reset no fue suficiente.")

        logger.warning(
            "Ejecutando Recuperación Nivel 2: Reiniciando navegador y reautenticando..."
        )
        try:
            context.close()
        except Exception:
            pass

        logger.info("Esperando unos segundos para que los servicios de SUNAT respondan...")
        time.sleep(BREAK_SECONDS_RECOVER)

        return self.Login_Sunat(playwright_instance)

    def Select_Cpe_Type(self, frame: FrameLocator, codigo_cpe: str):
        expected_text = MAP_TYPES_CPE.get(codigo_cpe, "Factura")
        btn_dropdown = (
            frame.get_by_role("button").filter(has_text=re.compile(r"^$")).first
        )
        if btn_dropdown.is_visible():
            btn_dropdown.click()
        else:
            frame.locator("div, span").filter(
                has_text=re.compile(r"^Seleccionar")
            ).first.click()

        option = frame.locator("div, li").filter(
            has_text=re.compile(rf"^{re.escape(expected_text)}$", re.IGNORECASE)
        ).first
        option.click()

    def Download_XML_File(self, page: Page, frame: FrameLocator):
        logger.info("Esperando que el modal 'Resultado' esté visible...")
        modal = frame.locator("div:has-text('Resultado'):visible").first
        modal.wait_for(state="visible", timeout=15000)

        page.wait_for_timeout(1000)

        candidate_xmls = frame.locator(
            "button[title*='XML' i], a[title*='XML' i], "
            "button[aria-label*='XML' i], a[aria-label*='XML' i], "
            "button[ngbtooltip*='XML' i], a[ngbtooltip*='XML' i]"
        )

        btn_xml = None

        if candidate_xmls.count() > 0:
            btn_xml = candidate_xmls.first
            logger.info(f"Botón XML localizado por atributo directo (coincidencias: {candidate_xmls.count()})")
        else:
            logger.warning("No se encontró por atributo directo. Buscando estructuralmente en la barra superior...")
            btn_pdf = modal.locator(
                "button:has(.fa-file-pdf), button[title*='PDF' i], a[title*='PDF' i]"
            ).first

            if btn_pdf.is_visible():
                btn_xml = btn_pdf.locator("xpath=following-sibling::*[1]")
                logger.info("Botón XML localizado como botón contiguo al PDF...")
            else:
                btns_bar = modal.locator("button:visible, a:visible")
                total_btns = btns_bar.count()
                logger.info(f"Botones detectados en el modal: {total_btns}...")
                if total_btns >= 2:
                    btn_xml = btns_bar.nth(1)
                    logger.info(
                        "Botón XML seleccionado por posición de barra (índice 1)..."
                    )

        if not btn_xml or not btn_xml.is_visible():
            raise Exception(
                "ERROR_UBICACION_BOTON: No se pudo localizar el botón XML en el visor."
            )

        try:
            btn_xml.highlight()
        except Exception:
            pass

        logger.info("Haciendo clic y esperando respuesta de SUNAT...")

        download_detect = []
        
        def capture_download(download):
            download_detect.append(download)

        page.on("download", capture_download)

        try:
            btn_xml.click()

            modal_error_frame = (
                frame.get_by_text("servicios de SUNAT")
                .or_(frame.get_by_text("reintentar en 5 minutos"))
                .first
            )
            modal_error_page = (
                page.get_by_text("servicios de SUNAT")
                .or_(page.get_by_text("reintentar en 5 minutos"))
                .first
            )

            ready = self.Wait_Condition(
                page=page,
                condition=lambda: (
                    bool(download_detect)
                    or modal_error_frame.is_visible()
                    or modal_error_page.is_visible()
                ),
                timeout_ms=25000,
                interval_ms=150
            )

            if not ready:
                raise Exception("TIMEOUT_DESCARGA_XML_SUNAT")

            if download_detect:
                download = download_detect[0]
                logger.info(f"Archivo .zip descargado con éxito de SUNAT: {download.suggested_filename}")
                return download

            if modal_error_frame.is_visible() or modal_error_page.is_visible():
                logger.error("Detectado 'Error del Servidor' de SUNAT al descargar XML. Abortando de inmediato...")
                raise Exception("SUNAT_ERROR_DEL_SERVIDOR")
        finally:
            page.remove_listener("download", capture_download)

    def Extract_XML_File(self, zip_path: Path) -> list[str]:
        extracted_xmls = []
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            for file_info in zip_ref.infolist():
                filename = Path(file_info.filename).name
                if filename.lower().endswith(".xml") and not filename.startswith("R-"):
                    destino = PATH_FOLDER_XMLS / filename
                    with (
                        zip_ref.open(file_info) as source,
                        open(destino, "wb") as target,
                    ):
                        shutil.copyfileobj(source, target)
                    extracted_xmls.append(filename)
        return extracted_xmls

    def Zip_Downloaded_XMLs(self, timestamp: str) -> Path | None:
        files_xml = list(PATH_FOLDER_XMLS.glob("*.xml"))
        if not files_xml:
            logger.warning("No se encontraron archivos XML para empaquetar en el ZIP de correo.")
            return None

        name_zip = f"XMLs_SUNAT_Descargados_{timestamp}.zip"
        path_zip_local = PATH_FOLDER_XMLS_ZIP / name_zip

        with zipfile.ZipFile(
            path_zip_local, "w", compression=zipfile.ZIP_DEFLATED
        ) as zip_out:
            for xml_file in files_xml:
                zip_out.write(xml_file, arcname=xml_file.name)

        size_mb = path_zip_local.stat().st_size / (1024 * 1024)
        logger.info(
            f"ZIP generado exitosamente: {name_zip} ({size_mb:.2f} MB con {len(files_xml)} comprobantes)."
        )
        return path_zip_local

    def Execute_Robot(self):
        global START_TIME

        if not self.Acquire_Lock():
            logger.warning("Ya existe una ejecución del robot en curso. Omitiendo proceso...")
            remove_log()
            sys.exit(0)

        for item in PATH_FOLDER_INPUTS_CLOUD.iterdir():
            dst_item_history = PATH_FOLDER_HISTORY_CLOUD / item.name
            dst_item_input = PATH_FOLDER_INPUTS / item.name

            if not dst_item_history.exists():
                shutil.copy2(item, dst_item_history)

            if not dst_item_input.exists():
                shutil.move(item, dst_item_input)

        try:
            START_TIME = get_current_datetime()
            name_process = "Descarga Automatizada de XMLs SUNAT"
            console.rule(f"[grey66]Proceso Iniciado: [bold white]{name_process}[/bold white][/grey66]")

            logger.info("Iniciando Robot...")

            control_storage_day(
                [
                    PATH_FOLDER_HISTORY_CLOUD,
                    PATH_GENERAL_CLOUD / "Xmls",
                    PATH_FOLDER_XMLS_ZIP_CLOUD,
                    PATH_FOLDER_PROCESSED_CLOUD,
                    PATH_FOLDER_QUARANTINE_CLOUD
                ]
            )

            remove_log()
            start_log(PATH_LOGS, FILE_LOG_EXISTS, True)

            df_validos, df_observados, archivos_origen = self.Load_Clean_Data()

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_observados.empty:
                ruta_obs = (
                    PATH_FOLDER_QUARANTINE / f"Comprobantes_Observados_{timestamp_actual}.xlsx"
                )
                export_df_to_excel(ruta_obs,df_observados[COLUMNS_FINAL])
                logger.warning(f"Se apartó {len(df_observados)} registros observados en: {ruta_obs.name}...")

            if df_validos.empty:
                raise Exception(f"No se encontraron registros válidos para procesar.\nUbicación Carpeta: {PATH_FOLDER_INPUTS}")

            total_rows = len(df_validos)

            df_lote_actual = df_validos.iloc[:MAX_RECORDS_PER_RUN].copy()

            df_remanentes_cola = df_validos.iloc[MAX_RECORDS_PER_RUN:].copy()
            df_remanentes_cola["DETALLE_ESTADO"] = "NO PROCESADO - REMANENTE"
            df_remanentes_cola["FECHA_PROCESO"] = None

            total_to_process = len(df_lote_actual)
            logger.info(
                f"Total disponible: {total_rows} registros. Cuota asignada para este ciclo: {total_to_process} registros."
            )

            map_xmls_locals = self.Index_XMLs_Locals()

            sunat_colapse = False

            with sync_playwright() as p:
                try:
                    context, page, frame = self.Login_Sunat(p)
                except SunatConexionError as sce:
                    logger.error(
                        f"Se presentó problemas al iniciar sesión en el portal SUNAT.\n{sce}"
                    )
                    context, page, frame = None, None, None, None
                    sunat_colapse = True
                
                contador_batch = 0

                if not sunat_colapse:
                    for idx, fila in df_lote_actual.iterrows():
                        contador_batch += 1
                        filtro = (
                            str(fila.get("filtro", "Recibido")).strip().capitalize()
                            or "Recibido"
                        )

                        ruc = str(fila["RUC"]).strip()
                        tipo = str(fila["COD_DCMNTO"]).strip()
                        serie = str(fila["SERIEDOC"]).strip()
                        numero = str(fila["NUMERODOC"]).strip()

                        num_limpio = (
                            str(int(numero)) if numero.isdigit() else numero
                        )

                        identificador_cpe = f"{serie}-{numero} (RUC: {ruc})"

                        archivo_existente = map_xmls_locals.get((serie, num_limpio))
                        if archivo_existente:
                            logger.info(
                                f"[{contador_batch}/{total_rows}] OMITIENDO: {identificador_cpe} ya existe en disco ({archivo_existente})."
                            )
                            df_lote_actual.at[idx, "ESTADO"] = "DESCARGADO"
                            df_lote_actual.at[idx, "DETALLE_ESTADO"] = (
                                f"RECUPERADO_LOCAL_PREVIO: {archivo_existente}"
                            )
                            df_lote_actual.at[idx, "FECHA_PROCESO"] = (
                                pd.Timestamp.now()
                            )
                            continue

                        logger.info(
                            f"[{contador_batch}/{total_rows}] Consultando SUNAT: {identificador_cpe}"
                        )

                        descarga_exitosa = False
                        detalle_resultado = "ERROR_DESCONOCIDO"

                        for intento in range(1, MAX_RETRIES_PER_RECORD + 1):
                            try:
                                frame.get_by_text(filtro, exact=True).first.click(timeout=10000)

                                input_ruc = frame.locator('input[name="rucEmisor"]')
                                input_ruc.fill("")
                                input_ruc.press_sequentially(ruc, delay=5)

                                self.Select_Cpe_Type(frame, tipo)

                                input_serie = frame.locator('input[name="serieComprobante"]')
                                input_serie.fill("")
                                input_serie.press_sequentially(serie, delay=15)

                                input_num = frame.locator('input[name="numeroComprobante"]')
                                input_num.fill("")
                                input_num.press_sequentially(numero, delay=15)

                                page.wait_for_timeout(300)

                                frame.get_by_role(
                                    "button", name=re.compile(r"Consultar")
                                ).click()

                                modal_resultado = frame.locator(
                                    "div:has-text('Resultado'):visible"
                                ).first 

                                pattern_error_sunat = re.compile(
                                    r"servicios de SUNAT|reintentar en 5 minutos", re.IGNORECASE
                                )

                                modal_error_page = page.get_by_text(pattern_error_sunat).first
                                modal_error_frame = frame.get_by_text(pattern_error_sunat).first

                                modal_error_page_aux = frame.get_by_text(
                                    re.compile(r"no puede completar|solicitud dentro del marco", re.IGNORECASE)
                                ).first

                                alerta_cpe_no_existe = frame.get_by_text(
                                    re.compile(r"No hay resultados|consulta realizada", re.IGNORECASE)
                                ).first

                                alerta_mayor_24 = frame.get_by_text(
                                    re.compile(r"mayor a 24 meses|pedido de descarga", re.IGNORECASE)
                                ).first

                                respuesta_servidor = self.Wait_Condition(
                                    page=page,
                                    condition=lambda: (
                                        modal_resultado.is_visible()
                                        or alerta_cpe_no_existe.is_visible()
                                        or alerta_mayor_24.is_visible()
                                        or modal_error_frame.is_visible()
                                        or modal_error_page.is_visible()
                                        or modal_error_page_aux.is_visible()
                                    ),
                                    timeout_ms=20000,
                                    interval_ms=150
                                )
                    
                                if not respuesta_servidor:
                                    raise Exception("SUNAT_SPINNER_TIMEOUT_BLOQUEADO")

                                if modal_error_frame.is_visible() or modal_error_page.is_visible() or modal_error_page_aux.is_visible():
                                    logger.error(
                                        f"Detectado 'Error del Servidor' de SUNAT en intento {intento}..."
                                    )
                                    raise Exception("SUNAT_ERROR_DEL_SERVIDOR")

                                if alerta_cpe_no_existe.is_visible():
                                    texto_alerta = alerta_cpe_no_existe.inner_text()

                                    logger.warning(f"SUNAT respondió: {texto_alerta}")
                                    
                                    detalle_resultado = f"COMPROBANTE_NO_EXISTE: {texto_alerta}"

                                    btn_ok = frame.get_by_role(
                                        "button", name="Aceptar"
                                    ).first

                                    try:
                                        btn_ok.click(timeout=4000)
                                    except Exception:
                                        btn_ok.click(force=True)

                                    frame.get_by_role("button", name="Limpiar").click()
                                    break

                                if alerta_mayor_24.is_visible():
                                    texto_alerta = alerta_mayor_24.inner_text()                       
                                    if not re.search(r"\d{8,}", texto_alerta):
                                        texto_alerta = alerta_mayor_24.locator(
                                            ".."
                                        ).inner_text()

                                    logger.warning(f"SUNAT respondió: {texto_alerta}")

                                    match_ticket = re.search(r"(\d{8,})", texto_alerta)
                                    ticket_id = (
                                        match_ticket.group(1)
                                        if match_ticket
                                        else "SIN_TICKET"
                                    )
                                    detalle_resultado = (
                                        f"HISTORICO_24_MESES_PEDIDO_DESCARGA_NRO_{ticket_id}"
                                    )
                                    logger.warning(f"Comprobante > 24 meses. N° Pedido: {ticket_id}")

                                    btn_ok = frame.get_by_role(
                                        "button", name="Aceptar"
                                    ).first
                                    try:
                                        btn_ok.click(timeout=4000)
                                    except Exception:
                                        btn_ok.click(force=True)

                                    frame.get_by_role("button", name="Limpiar").click()
                                    break

                                download = self.Download_XML_File(page, frame)

                                temp_zip_path = PATH_FOLDER_XMLS / download.suggested_filename
                                download.save_as(temp_zip_path)

                                archivos_ok = self.Extract_XML_File(temp_zip_path)
                                temp_zip_path.unlink(missing_ok=True)

                                if archivos_ok:
                                    descarga_exitosa = True
                                    detalle_resultado = f"DESCARGA_EXITOSA: {archivos_ok[0]}"
                                    logger.info(f"Descarga y extracción XML exitosa: {archivos_ok[0]}")
                                else:
                                    detalle_resultado = "ZIP_SIN_XML_VALIDO"

                                frame.get_by_role("button", name="Close").click()
                                frame.get_by_role("button", name="Limpiar").click()
                                break
                            except Exception as ex:
                                logger.warning(
                                    f"Fallo en intento {intento}/{MAX_RETRIES_PER_RECORD} para {identificador_cpe}: {ex}..."
                                )
                                detalle_resultado = f"ERROR_INTENTO_{intento}: {str(ex)[:35]}"

                                is_error_server = "SUNAT_ERROR_DEL_SERVIDOR" in str(ex)

                                try:
                                    if is_error_server:
                                        logger.warning("Fallo crítico de servidor. Saltando directo a reinicio de navegador (Nivel 2)...")
                                        try:
                                            context.close()
                                        except Exception:
                                            pass
                                        context, page, frame = self.Login_Sunat(p)
                                    else:
                                        context, page, frame = self.Recover_Screen_Restart(p, context, page)
                                except SunatConexionError as sce:
                                    logger.error(
                                        f"SUNAT colapsó. Deteniendo lote: {sce}"
                                    )
                                    sunat_colapse = True
                                    break

                        df_lote_actual.at[idx, "ESTADO"] = (
                            "DESCARGADO" if descarga_exitosa else "NO_DESCARGADO"
                        )
                        df_lote_actual.at[idx, "DETALLE_ESTADO"] = detalle_resultado
                        df_lote_actual.at[idx, "FECHA_PROCESO"] = pd.Timestamp.now()

                        if sunat_colapse:
                            break
                        
                        self.Random_Delay()

                        if (contador_batch % BATCH_SIZE == 0) and (contador_batch < total_rows and contador_batch < MAX_RECORDS_PER_RUN):
                            logger.info(
                                f"Lote de {BATCH_SIZE} alcanzado. Pausando {BREAK_MINUTE_BATCH} min.."
                            )
                            time.sleep(BREAK_MINUTE_BATCH * 60)

                    if page and not page.is_closed():
                        try:
                            logger.info("Cerrando sesión en SUNAT...")
                            page.on("dialog", lambda dialog: dialog.accept())
                            page.get_by_role("button", name="Salir").click()
                            page.wait_for_url(
                                "**/loginMenuSol*",
                                timeout=10000,
                                wait_until="commit",
                            )
                            logger.info("Sesión cerrada correctamente...")
                        except Exception as e:
                            logger.warning(f"Aviso durante el cierre de sesión: {e}")
                            pass

                if context:
                    try:
                        context.close()
                    except Exception:
                        pass

            df_descargados = df_lote_actual[df_lote_actual["ESTADO"] == "DESCARGADO"].copy()

            observados_sunat = df_lote_actual["DETALLE_ESTADO"].str.contains(
                r"COMPROBANTE_NO_EXISTE|HISTORICO_24_MESES", regex=True
            )
            df_observados_sunat = df_lote_actual[observados_sunat].copy()

            if not df_observados_sunat.empty:
                df_observados_sunat["ESTADO"] = "OBSERVADO_SUNAT"

            df_fallidos_tecnicos = df_lote_actual[
                ((df_lote_actual["ESTADO"] == "NO_DESCARGADO") | (df_lote_actual["ESTADO"] == "PENDIENTE"))
                & (~observados_sunat)
            ].copy()

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_observados_sunat.empty:
                ruta_obs_sunat = (
                    PATH_FOLDER_QUARANTINE
                    / f"Comprobantes_No_Descargados_SUNAT_{timestamp_actual}.xlsx"
                )
                export_df_to_excel(ruta_obs_sunat,df_observados_sunat[COLUMNS_FINAL])

                copy_items_to_folder(PATH_FOLDER_QUARANTINE,PATH_FOLDER_QUARANTINE_CLOUD,True,False)

                logger.warning(
                    f"Se guardó Excel con registros observados por SUNAT en OneDrive: {PATH_FOLDER_QUARANTINE_CLOUD.parent}/{PATH_FOLDER_QUARANTINE_CLOUD.name} ({len(df_observados_sunat)} filas)"
                )

            for archivo_leido in archivos_origen:
                if archivo_leido.exists():
                    archivo_leido.unlink()
                    logger.info(f"Se eliminó archivo procesado: {archivo_leido.name}")

            df_nueva_cola = pd.concat(
                [df_fallidos_tecnicos, df_remanentes_cola], ignore_index=True
            )

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_nueva_cola.empty:
                ruta_cola_inputs = PATH_FOLDER_INPUTS / f"Comprobantes_Pendientes_{timestamp_actual}.xlsx"
                df_nueva_cola["ESTADO"] = "PENDIENTE"
                export_df_to_excel(ruta_cola_inputs,df_nueva_cola[COLUMNS_FINAL])

                logger.warning(
                    f"Quedan {len(df_nueva_cola)} comprobantes pendientes ({len(df_fallidos_tecnicos)} por reintentar + {len(df_remanentes_cola)} fuera de cuota)."
                    f"Actualizado en: {ruta_cola_inputs.name} para la próxima ejecución."
                )
            else:
                logger.info("Se descargó los XMLs de todos los comprobantes. Carpeta Inputs limpia...")

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_descargados.empty:
                ruta_proc = PATH_FOLDER_PROCESSED / f"Comprobantes_Descargados_{timestamp_actual}.xlsx"
                export_df_to_excel(ruta_proc,df_descargados[COLUMNS_FINAL])

                copy_items_to_folder(PATH_FOLDER_PROCESSED,PATH_FOLDER_PROCESSED_CLOUD,True,False)

                logger.info(
                    f"Se guardó Excel con registros procesados a OneDrive: {PATH_FOLDER_PROCESSED_CLOUD.parent}/{PATH_FOLDER_PROCESSED_CLOUD.name} ({len(df_descargados)} filas)"
                )

                copy_items_to_folder(PATH_FOLDER_XMLS,PATH_FOLDER_XMLS_CLOUD,False,True)
                        
                logger.info(
                    f"Se guardó Lote de XMLs descargados a OneDrive: {PATH_FOLDER_XMLS_CLOUD.parent}/{PATH_FOLDER_XMLS_CLOUD.name} ({len(df_descargados)} archivos)"
                )

                PATH_FOLDER_XMLS_ZIP.mkdir(parents=True, exist_ok=True)

                zip_generado = self.Zip_Downloaded_XMLs(timestamp_actual)

                if zip_generado and zip_generado.exists():
                    copy_items_to_folder(PATH_FOLDER_XMLS_ZIP, PATH_FOLDER_XMLS_ZIP_CLOUD, True, False)
                    logger.info(
                        f"Se guardó Consolidado ZIP en OneDrive: {PATH_FOLDER_XMLS_ZIP_CLOUD}/{zip_generado.name}"
                    )

            shutil.rmtree(PATH_FOLDER_XMLS.parent, ignore_errors=True)

            self.Finish_Process(True)
        except Exception as e:
            self.Finish_Process(False, str(e))
        finally:
            self.Release_Lock()

    def Finish_Process(self, is_success: bool = False, msg_error: str|None = None):
        global END_TIME, FILE_LOG_EXISTS

        END_TIME = get_current_datetime()

        if is_success:
            logger.success("Ejecución exitosa: Se procesó la información.")
            close_rule = lambda: console.rule("[grey66]Proceso Finalizado[/grey66]")
            exit_code = 0
        else:
            logger.error(f"Proceso Incompleto. Detalle: {msg_error}")
            close_rule = lambda: show_custom_rule(console, "Proceso Finalizado con Error", state="Error")
            exit_code = 1

        total_seconds = int((END_TIME - START_TIME).total_seconds())
        minutes, seconds = divmod(total_seconds, 60)
        difference_formatted = f"{minutes} minuto(s), {seconds} segundo(s)"
        
        remove_log()
        FILE_LOG_EXISTS = add_log_file(PATH_LOGS, FILE_LOG_EXISTS, True)
        logger.info(f"Tiempo de proceso: {difference_formatted}")
        add_log_console()
        print(f"[dark_orange]Tiempo de proceso: {difference_formatted}[/dark_orange]")

        close_rule()
        sys.exit(exit_code)

if __name__ == "__main__":
    FILE_LOG_EXISTS = start_log(PATH_LOGS, FILE_LOG_EXISTS, False)
    console = Console()
    Robot_Sunat()