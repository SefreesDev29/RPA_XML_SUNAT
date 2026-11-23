import multiprocessing
import re
from pathlib import Path
import os
import time
import io
import json
import zipfile
from typing import Callable
from rich.console import Console
from Config import settings, BASE_DIR
from loguru import logger
from playwright.sync_api import (
    sync_playwright,
    Playwright,
    Page,
    BrowserContext,
    FrameLocator,
)

from Utils import (
    print,
    add_log_console,
    add_log_file,
    send_email_graph,
    build_sunat_email_html,
    export_df_to_excel,
    get_current_datetime,
    get_random_delay,
    copy_items_to_folder,
    control_storage_day,
    get_coord_screen_x,
    remove_temporal_folder,
    datetime,
    pd,
    remove_log,
    show_custom_rule,
    start_log,
    is_pid_running,
    sys,
    shutil,
    ctypes
)

PATH_GENERAL = BASE_DIR
PATH_GENERAL_CLOUD = Path(
    settings.PATH_CLOUD
)

AVAILABLE_AREAS = ["Siniestros", "Salud"]

SUNAT_CREDENTIALS = {
    "LPG": {
        "ruc": settings.SUNAT_RUC_LPG,
        "user": settings.SUNAT_USER_PROD_LPG,
        "password": settings.SUNAT_PASSWORD_PROD_LPG,
    },
    "LPE": {
        "ruc": settings.SUNAT_RUC_LPE,
        "user": settings.SUNAT_USER_PROD_LPE,
        "password": settings.SUNAT_PASSWORD_PROD_LPE,
    },
}

RUC_INVALID_SUNAT_SET = {settings.SUNAT_RUC_LPG, settings.SUNAT_RUC_LPE}

TENANT_ID = settings.TENANT_ID
CLIENT_ID = settings.CLIENT_ID
CLIENT_SECRET = settings.CLIENT_SECRET
MAIL_BOX = settings.MAIL_BOX
RECIPIENTS_MAIL_LIST = settings.RECIPIENTS_MAIL_LIST

BROWSER_MINIMIZED = settings.BROWSER_MINIMIZED

URL_LOGIN_SUNAT = (
    "https://api-seguridad.sunat.gob.pe/v1/clientessol/4f3b88b3-d9d6-402a-b85d-6a0bc857746a/oauth2/loginMenuSol"
    "?lang=es-PE&showDni=true&showLanguages=false"
    "&originalUrl=https://e-menu.sunat.gob.pe/cl-ti-itmenu/AutenticaMenuInternet.htm"
    "&state=rO0ABXNyABFqYXZhLnV0aWwuSGFzaE1hcAUH2sHDFmDRAwACRgAKbG9hZEZhY3RvckkACXRocmVzaG9sZHhwP0AAAAAAAAx3CAAAABAAAAADdAADZXhlcHQABnBhcmFtc3QASyomKiYvY2wtdGktaXRtZW51L01lbnVJbnRlcm5ldC5odG0mYjY0ZDI2YThiNWFmMDkxOTIzYjIzYjY0MDdhMWMxZGI0MWU3MzNhNnQABGV4ZWNweA=="
)

QUOTA_MORNING = settings.QUOTA_MORNING
QUOTA_AFTERNOON = settings.QUOTA_AFTERNOON
QUOTA_NIGHT = settings.QUOTA_NIGHT
MAX_DAILY_QUOTA = settings.MAX_DAILY_QUOTA
MAX_ZIP_SIZE_MB = settings.MAX_ZIP_SIZE_MB

BREAK_MIN_CALL = 2.0
BREAK_MAX_CALL = 2.5
BATCH_SIZE = 500
BREAK_SECONDS_RECOVER = 3
BREAK_MINUTE_BATCH = 1.0
MAX_RETRIES_PER_RECORD = 3

MAP_TYPES_CPE = {
    "01": "Factura",
    "03": "Boleta de Venta",
    "07": "Factura - Nota de Crédito",
    "08": "Factura - Nota de Débito",
    "02": "Recibo por Honorarios",
}

COLUMNS_FINAL = ["ARCHIVO_ORIGEN", "EMPRESA", "#EXP", "RUC", "COD_DCMNTO", "TPO_DCMNTO", "SERIEDOC", "NUMERODOC", "RAZON_SOCIAL", "ESTADO", "DETALLE_ESTADO", "FECHA_PROCESO"]

class SunatConexionError(Exception):
    pass

class Robot_Sunat:
    def __init__(self, area: str):
        self.area = area.strip()
        self.console = Console()
        self.start_time = get_current_datetime()
        self.offscreen_x = get_coord_screen_x()
        self.end_time = self.start_time
        self.period_time = f"{self.start_time.year}{self.start_time.month:02d}{self.start_time.day:02d}"

        self.path_area_cloud = PATH_GENERAL_CLOUD / self.area / "XMLs_SUNAT"
        self.path_folder_history_cloud = self.path_area_cloud / "Historicos"
        self.path_folder_inputs_cloud = self.path_area_cloud / "Inputs"
        self.path_folder_xmls_cloud = self.path_area_cloud / "Xmls" / f"Proceso_{self.period_time}"
        self.path_folder_xmls_zip_cloud = self.path_area_cloud / "Comprimidos_Correo"
        self.path_folder_processed_cloud = self.path_area_cloud / "Procesados"
        self.path_folder_quarantine_cloud = self.path_area_cloud / "Observados"

        self.path_area_local = PATH_GENERAL / "Areas_Local" / self.area
        self.path_folder_inputs = self.path_area_local / "Inputs"
        self.path_folder_processed = self.path_area_local / "Outputs" / "Procesados"
        self.path_folder_quarantine = self.path_area_local / "Outputs" / "Observados"
        self.path_folder_xmls = self.path_area_local / "Outputs" / "Xmls" / f"Proceso_{self.period_time}"
        self.path_folder_xmls_zip = self.path_area_local / "Outputs" / "Xmls" / f"Comprimidos_{self.period_time}"

        self.path_perfil_edge = self.path_area_local / "Perfil_Edge"
        self.path_lock_file = self.path_area_local / f"Robot_{self.area}.lock"
        self.path_quota_file = self.path_area_local / f"Cuota_Diaria_{self.period_time}.json"
        self.path_onedrive_lock = self.path_area_local / f"onedrive_cleanup_{self.area}.lock"

        self.path_folder_logs = PATH_GENERAL / "Logs"
        self.path_logs = self.path_folder_logs / f"Log_Robot_Sunat_{self.area}_{self.period_time}.log"
        self.file_log_exists = False

        self.Execute_Robot()

    def Write_Lock_Atomic(self) -> bool:
        content = f"{os.getpid()}|{time.time()}".encode("utf-8")
        try:
            fd = os.open(
                str(self.path_lock_file),
                os.O_CREAT | os.O_EXCL | os.O_WRONLY,
            )
            try:
                os.write(fd, content)
            finally:
                os.close(fd)
            return True
        except FileExistsError:
            return False

    def Acquire_Lock(self) -> bool:
        self.path_area_local.mkdir(parents=True, exist_ok=True)

        if self.Write_Lock_Atomic():
            return True

        try:
            contenido_raw = self.path_lock_file.read_text(
                encoding="utf-8"
            ).strip()
            pid_guardado = int(contenido_raw.split("|")[0])
            tiempo_transcurrido = (
                time.time() - self.path_lock_file.stat().st_mtime
            )

            proceso_vivo = is_pid_running(pid_guardado)

            if proceso_vivo and tiempo_transcurrido <= 12600:
                logger.warning(
                    f"[{self.area}] Bloqueo activo por PID {pid_guardado} "
                    f"(en ejecución hace {int(tiempo_transcurrido // 60)} min)."
                )
                return False

            motivo = (
                f"PID {pid_guardado} ya no existe en Windows"
                if not proceso_vivo
                else "superó el tiempo máximo de 3.5 horas"
            )
            logger.warning(
                f"[{self.area}] Detectado lock huérfano ({motivo}). Liberando..."
            )
            self.path_lock_file.unlink(missing_ok=True)

        except Exception as e:
            logger.warning(
                f"[{self.area}] Lock ilegible ({e}). Reemplazando..."
            )
            self.path_lock_file.unlink(missing_ok=True)

        return self.Write_Lock_Atomic()

    def Release_Lock(self) -> None:
        self.path_lock_file.unlink(missing_ok=True)

    def Ingest_Cloud_Inputs(self):
        if not self.path_folder_inputs_cloud.exists():
            logger.warning(
                f"[{self.area}] No existe la carpeta Inputs en OneDrive: {self.path_folder_inputs_cloud}"
            )
            return

        self.path_folder_history_cloud.mkdir(parents=True, exist_ok=True)

        for item in self.path_folder_inputs_cloud.iterdir():
            if not item.is_file() or item.suffix.lower() != ".xlsx":
                continue
            dst_item_history = self.path_folder_history_cloud / item.name
            dst_item_input = self.path_folder_inputs / item.name

            if not dst_item_history.exists():
                shutil.copy2(item, dst_item_history)

            if not dst_item_input.exists():
                shutil.move(item, dst_item_input)

    def Index_XMLs_Locals(self) -> dict[tuple[str, str], str]:  
        indice_xmls: dict[tuple[str, str], str] = {}
        if not self.path_folder_xmls.exists():
            return indice_xmls

        patron_estandar = re.compile(r"^(\d{11})-\d{2}-([A-Za-z0-9]{4})-(\d+)\.xml$", re.IGNORECASE)
        patron_cpe = re.compile(r"([A-Za-z0-9]{4})-(\d+)\.xml$", re.IGNORECASE)

        for archivo in self.path_folder_xmls.glob("*.xml"):
            nombre = archivo.name

            match_std = patron_estandar.match(nombre)
            if match_std:
                ruc_archivo = match_std.group(1)
                serie_archivo = self.Normalize_Serie(match_std.group(2))
                numero_archivo = str(int(match_std.group(3)))
                indice_xmls[(ruc_archivo, serie_archivo, numero_archivo)] = (
                    nombre
                )
                continue

            match = patron_cpe.search(nombre)
            if match:
                serie_archivo = self.Normalize_Serie(
                    match.group(1).upper().strip()
                )
                numero_raw = match.group(2).strip()
                ruc_archivo = ""

                if len(numero_raw) > 11:
                    ruc_archivo = numero_raw[-11:]
                    numero_raw = numero_raw[:-11]

                numero_archivo = (
                    str(int(numero_raw)) if numero_raw.isdigit() else numero_raw
                )
                indice_xmls[(ruc_archivo, serie_archivo, numero_archivo)] = nombre

        logger.info(
            f"[{self.area}] Índice local cargado: {len(indice_xmls)} XMLs encontrados previamente en disco."
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
        if not self.path_folder_inputs.exists():
            raise Exception(f"[{self.area}] No se encontró carpeta a procesar: {self.path_folder_inputs}")

        excel_files = list(self.path_folder_inputs.glob("*.xlsx"))

        if not excel_files:
            raise Exception(f"[{self.area}] No se encontraron archivos Excel para procesar.\nUbicación Carpeta Cloud: {self.path_folder_inputs_cloud}\nUbicación Carpeta Local: {self.path_folder_inputs}")

        dfs: list[pd.DataFrame] = []
        company_default = "LPG" if self.area.upper() == "SINIESTROS" else "LPE"

        for file in excel_files:
            try:
                temp_df = pd.read_excel(file, dtype=str)
                temp_df.columns = [c.strip().upper() for c in temp_df.columns]

                if "ARCHIVO_ORIGEN" not in temp_df.columns:
                    temp_df["ARCHIVO_ORIGEN"] = file.name

                if "EMPRESA" not in temp_df.columns:
                    temp_df["EMPRESA"] = company_default

                for col in ["EMPRESA", "#EXP", "RUC", "COD_DCMNTO", "TPO_DCMNTO", "SERIEDOC", "NUMERODOC", "RAZON_SOCIAL"]:
                    if col not in temp_df.columns:
                        temp_df[col] = ""
                    temp_df[col] = (
                        temp_df[col].fillna("").astype(str).str.strip()
                    )

                temp_df["EMPRESA"] = (
                    temp_df["EMPRESA"]
                    .str.upper()
                    .replace({"LPS": "LPE", "": company_default})
                )

                dfs.append(temp_df)
                logger.info(
                    f"[{self.area}] Cargado archivo: {file.name} ({len(temp_df)} filas)"
                )
            except Exception as e:
                logger.error(
                    f"[{self.area}] Error al leer Excel. Ubicación Excel: '{file.parent}/{file.name}'.\n{e}"
                )

        if not dfs:
            raise Exception(
                f"[{self.area}] No se pudieron consolidar archivos Excel en: {self.path_folder_inputs}"
            )

        df_consolidado = pd.concat(dfs, ignore_index=True)


        df_consolidado["RUC"] = df_consolidado["RUC"].str.replace(" ", "", regex=False).str.replace(",", "", regex=False).str.split(".").str[0]
        df_consolidado["SERIEDOC"] = df_consolidado["SERIEDOC"].apply(self.Normalize_Serie)
        df_consolidado["COD_DCMNTO"] = df_consolidado["COD_DCMNTO"].apply(
            lambda x: str(x).zfill(2) if str(x).isdigit() and len(str(x)) == 1 else x
        )

        rows_preview = len(df_consolidado)
        df_consolidado = df_consolidado.drop_duplicates(
            subset=["EMPRESA", "RUC", "COD_DCMNTO", "SERIEDOC", "NUMERODOC"]
        ).copy()
        if len(df_consolidado) < rows_preview:
            logger.info(
                f"[{self.area}] Se descartaron {rows_preview - len(df_consolidado)} registros duplicados."
            )

        is_company_valid = df_consolidado["EMPRESA"].isin(["LPG", "LPE"])
        is_ruc_valid = df_consolidado["RUC"].str.match(r"^\d{11}$")
        fields_complete = (
            (~df_consolidado["RUC"].isin(RUC_INVALID_SUNAT_SET)) &
            (df_consolidado["SERIEDOC"] != "")
            & (df_consolidado["SERIEDOC"].str.len() > 3)
            & (df_consolidado["NUMERODOC"] != "")
        )
        mask_valid = is_company_valid & is_ruc_valid & fields_complete

        df_validos = df_consolidado[mask_valid].copy()
        df_observados = df_consolidado[~mask_valid].copy()

        df_validos["ESTADO"] = "PENDIENTE"
        df_validos["DETALLE_ESTADO"] = ""
        df_validos["FECHA_PROCESO"] = pd.Timestamp.now()

        df_observados["ESTADO"] = "OBSERVADO"
        if df_observados.empty:
            df_observados["DETALLE_ESTADO"] = ""
        else:
            df_observados["DETALLE_ESTADO"] = df_observados.apply(
                lambda row: (
                    "EMPRESA_VACÍA_O_INVÁLIDA"
                    if row["EMPRESA"] not in ("LPG", "LPE")
                    else "RUC_LA_POSITIVA_SE_DESCARTA"
                    if (row["RUC"] in RUC_INVALID_SUNAT_SET)
                    else "RUC_VACÍO_O_INVÁLIDO"
                    if not re.match(r"^\d{11}$", row["RUC"])
                    else "SERIE_VACÍO_O_INVÁLIDO"
                    if (row["SERIEDOC"] == "" or len(row["SERIEDOC"]) <= 3)
                    else "NÚMERO_VACÍO_O_INVÁLIDO"
                ),
                axis=1,
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

    def Move_Browser_Offscreen(self, page: Page):
        try:
            unique_title = f"EDGE_RPA_{self.area}_{os.getpid()}"
            page.evaluate(f"document.title = '{unique_title}'")

            user32 = ctypes.WinDLL("user32", use_last_error=True)
            HWND_BOTTOM = 1
            SWP_NOSIZE = 0x0001
            SWP_NOACTIVATE = 0x0010
            SWP_SHOWWINDOW = 0x0040

            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
            target_hwnd = None

            def enum_windows_cb(hwnd, _lparam):
                nonlocal target_hwnd
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buff, length + 1)
                    if unique_title in buff.value:
                        target_hwnd = hwnd
                        return False
                return True

            user32.EnumWindows(WNDENUMPROC(enum_windows_cb), 0)

            if target_hwnd:
                user32.SetWindowPos(target_hwnd,HWND_BOTTOM,self.offscreen_x,0,0,0,SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,)
        except Exception as e:
            logger.debug(
                f"[{self.area}] No se pudo mover la ventana fuera de pantalla: {e}"
            )

    def Disable_Edge_Download_Popup(self, page: Page) -> None:
        target_prefs = {
            "browser.show_hub_popup_on_download_start": False,
            "download.prompt_for_download": False,
            "credentials_enable_service": False,
            "credentials_enable_autosignin": False,
            "credentials_enable_autofill_passwords": False,
            "credentials_enable_passwordgeneration": False,
            "credentials_enable_autosave": False,
        }
        try:
            page.goto("edge://settings/downloads", wait_until="domcontentloaded", timeout=15000)
            page.wait_for_function(
                "() => Boolean(window.chrome && chrome.settingsPrivate && chrome.settingsPrivate.getPref && chrome.settingsPrivate.setPref)",
                timeout=10000,
            )
            results = page.evaluate(
                """async (targetPrefs) => {
                    const callWithTimeout = (fn) => new Promise((resolve, reject) => {
                        const timer = setTimeout(() => reject(new Error("settingsPrivate callback timeout")), 5000);
                        try {
                            fn((res) => {
                                clearTimeout(timer);
                                resolve(res);
                            });
                        } catch (err) {
                            clearTimeout(timer);
                            reject(err);
                        }
                    });

                    const getPref = (key) => callWithTimeout((cb) => chrome.settingsPrivate.getPref(key, cb));
                    const setPref = (key, val) => callWithTimeout((cb) => chrome.settingsPrivate.setPref(key, val, "", cb));

                    const results = {};
                    for (const [key, desired] of Object.entries(targetPrefs)) {
                        const beforeObj = await getPref(key);
                        const beforeVal = beforeObj ? beforeObj.value : undefined;
                        let changed = false;
                        let setOk = true;
                        if (beforeVal !== desired) {
                            setOk = Boolean(await setPref(key, desired));
                            changed = true;
                        }
                        const afterObj = await getPref(key);
                        const afterVal = afterObj ? afterObj.value : undefined;
                        results[key] = {
                            before: beforeVal,
                            after: afterVal,
                            changed,
                            ok: Boolean(setOk && afterVal === desired),
                        };
                    }
                    return results;
                }""",
                target_prefs,
            )
            failed = {k: v for k, v in results.items() if not v.get("ok")}
            if failed:
                logger.warning(
                    f"[{self.area}] No se pudieron confirmar todas las preferencias de Edge: {failed}"
                )
            elif any(v.get("changed") for v in results.values()):
                page.wait_for_timeout(300)
                logger.info(
                    f"[{self.area}] Preferencias de Edge configuradas (popups de descargas y contraseñas desactivados)."
                )
            else:
                logger.debug(
                    f"[{self.area}] Preferencias de Edge ya estaban configuradas."
                )
        except Exception as e:
            logger.warning(
                f"[{self.area}] No se pudieron configurar las preferencias de Edge vía settingsPrivate: {e}"
            )

    def Login_Sunat(self, playwright_instance: Playwright, company: str, max_login_attempts: int = 3) -> tuple[BrowserContext, Page, FrameLocator]:
        creds = SUNAT_CREDENTIALS.get(company.upper())
        if not creds or not creds["ruc"] or not creds["user"] or not creds["password"]:
            raise SunatConexionError(
                f"[{self.area}] Credenciales incompletas para la empresa {company}."
            )

        for attempt in range(1, max_login_attempts + 1):
            context = None
            try:
                logger.info(f"[{self.area} | {company}] Lanzando Edge y autenticando en SOL (Intento {attempt}/{max_login_attempts})...")

                window_args = (
                    [
                        "--window-size=1366,768",
                        f"--window-position={self.offscreen_x},0",
                    ]
                    if BROWSER_MINIMIZED
                    else ["--start-maximized"]
                )

                context = playwright_instance.chromium.launch_persistent_context(
                    user_data_dir=str(self.path_perfil_edge),
                    channel="msedge",
                    headless=False,
                    chromium_sandbox=True,
                    ignore_default_args=["--no-sandbox"],
                    no_viewport=True,
                    accept_downloads=True,
                    args=[
                        *window_args,
                        "--disable-blink-features=AutomationControlled",
                        "--disable-backgrounding-occluded-windows",
                        "--disable-renderer-backgrounding",
                        "--disable-features=CalculateNativeWinOcclusion",
                        "--disable-background-mode",
                        "--no-first-run",
                        "--no-default-browser-check",
                    ],
                )

                context.clear_cookies()
                page = context.pages[0] if context.pages else context.new_page()

                if BROWSER_MINIMIZED:
                    self.Move_Browser_Offscreen(page)

                self.Disable_Edge_Download_Popup(page)

                page.goto(URL_LOGIN_SUNAT, wait_until="domcontentloaded", timeout=30000)

                field_ruc = page.get_by_role("textbox", name="RUC")
                field_ruc.wait_for(state="visible", timeout=30000)
                field_ruc.click()
                field_ruc.press_sequentially(creds["ruc"], delay=5)

                field_user = page.get_by_role("textbox", name="Usuario")
                field_user.click()
                field_user.press_sequentially(creds["user"], delay=15)

                field_pass = page.get_by_role("textbox", name="Contraseña")
                field_pass.click()
                field_pass.press_sequentially(creds["password"], delay=15)

                page.wait_for_timeout(300)

                page.get_by_role("button", name="Iniciar sesión").click()

                page.wait_for_url("**/MenuInternet.htm*", timeout=30000)
                logger.info(f"[{self.area} | {company}] Credenciales aceptadas por SUNAT...")

                page.wait_for_timeout(300)

                tree_menu = [
                    ("#divOpcionServicio1", "Módulo Empresas"),
                    ("#nivel1_11", "Comprobantes de Pago"),
                    ("#nivel2_11_38", "SEE - Facturas y Recibos"),
                    ("#nivel3_11_38_1", "Factura Electrónica"),
                    ("#nivel4_11_38_1_1_1", "Nueva Consulta de Comprobantes"),
                ]

                for selector, fallback_text in tree_menu:
                    nodo = page.locator(selector)
                    try:
                        nodo.wait_for(state="visible", timeout=10000)
                    except Exception:
                        if fallback_text:
                            logger.warning(
                                f"[{self.area} | {company}] Selector {selector} no visible, usando fallback de texto '{fallback_text}'."
                            )
                            nodo = page.get_by_text(fallback_text).first
                            nodo.wait_for(state="visible", timeout=10000)
                        else:
                            raise
                    nodo.click()
                    page.wait_for_timeout(100)

                frame = page.frame_locator('iframe[name="iframeApplication"]')

                logger.info(f"[{self.area} | {company}] Esperando carga del formulario de consulta dentro del iframe...")

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
                    interval_ms=100,
                )

                if not ready:
                    raise Exception("TIMEOUT_CARGA_IFRAME_FORMULARIO")

                if modal_error_frame.is_visible() or modal_error_page.is_visible():
                    logger.error(f"[{self.area} | {company}] Detectado 'Error del Servidor' de SUNAT en intento {attempt}...")
                    raise Exception(
                        "SUNAT_CAIDO_AL_INICIAR: Error del Servidor detectado."
                    )

                logger.info(
                    f"[{self.area} | {company}] Formulario cargado y listo para operar..."
                )
                return context, page, frame
            except Exception as e:
                logger.warning(
                    f"[{self.area} | {company}] Fallo en login {attempt}/{max_login_attempts}: {e}"
                )
                if context:
                    try:
                        context.close()
                    except Exception:
                        pass

                if attempt < max_login_attempts:
                    logger.info(f"[{self.area} | {company}] Esperando unos segundos para que los servicios de SUNAT respondan...")
                    time.sleep(BREAK_SECONDS_RECOVER)
                else:
                    raise SunatConexionError(
                        f"SUNAT inaccesible para {company} tras {max_login_attempts} intentos."
                    )

    def Recover_Screen_Restart(self, playwright_instance: Playwright, context: BrowserContext, page: Page, company: str) -> tuple[BrowserContext, Page, FrameLocator]:
        logger.warning(f"[{self.area} | {company}] Iniciando protocolo de recuperación de pantalla...")

        try:
            btn_new_search = page.locator("#nivel4_11_38_1_1_1").or_(
                page.get_by_text("Nueva Consulta de")
            ).first
            if btn_new_search.is_visible():
                btn_new_search.click(timeout=5000)
                page.wait_for_timeout(1500)

                frame = page.frame_locator('iframe[name="iframeApplication"]')
                frame.get_by_text("Recibido").wait_for(
                    state="visible", timeout=15000
                )
                logger.info(
                    f"[{self.area} | {company}] Recuperación Nivel 1 (Soft Reset) exitosa..."
                )
                return context, page, frame
        except Exception:
            logger.debug("Soft Reset no fue suficiente.")

        logger.warning(
            f"[{self.area} | {company}] Ejecutando Recuperación Nivel 2: Reiniciando navegador..."
        )
        try:
            context.close()
        except Exception:
            pass

        logger.info(f"[{self.area} | {company}] Esperando unos segundos para que los servicios de SUNAT respondan...")
        time.sleep(BREAK_SECONDS_RECOVER)

        return self.Login_Sunat(playwright_instance, company)

    def Close_Sunat_Session(self, context: BrowserContext, page: Page, company: str):
        if page and not page.is_closed():
            try:
                logger.info(f"[{self.area} | {company}] Cerrando sesión en SUNAT...")
                page.on("dialog", lambda dialog: dialog.accept())
                page.get_by_role("button", name="Salir").click(timeout=5000)
                page.wait_for_url(
                    "**/loginMenuSol*", timeout=10000, wait_until="commit"
                )
            except Exception as e:
                logger.warning(f"[{self.area} | {company}] Aviso al cerrar sesión: {e}")

        if context:
            try:
                context.close()
            except Exception:
                pass

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
        modal = frame.locator("div:has-text('Resultado'):visible").first
        modal.wait_for(state="visible", timeout=15000)
        
        page.wait_for_timeout(1250)

        candidate_xmls = frame.locator(
            "button[title*='XML' i], a[title*='XML' i], "
            "button[aria-label*='XML' i], a[aria-label*='XML' i], "
            "button[ngbtooltip*='XML' i], a[ngbtooltip*='XML' i]"
        )

        btn_xml = None

        if candidate_xmls.count() > 0:
            btn_xml = candidate_xmls.first
            logger.info(f"[{self.area}] Botón XML localizado por atributo directo (coincidencias: {candidate_xmls.count()})")
        else:
            logger.warning(f"[{self.area}] No se encontró por atributo directo. Buscando estructuralmente en la barra superior...")
            btn_pdf = modal.locator(
                "button:has(.fa-file-pdf), button[title*='PDF' i], a[title*='PDF' i]"
            ).first

            if btn_pdf.is_visible():
                btn_xml = btn_pdf.locator("xpath=following-sibling::*[1]")
                logger.info(f"[{self.area}] Botón XML localizado como botón contiguo al PDF...")
            else:
                btns_bar = modal.locator("button:visible, a:visible")
                total_btns = btns_bar.count()
                logger.info(f"[{self.area}] Botones detectados en el modal: {total_btns}...")
                if total_btns >= 2:
                    btn_xml = btns_bar.nth(1)
                    logger.info(f"[{self.area}] Botón XML seleccionado por posición de barra (índice 1)...")

        if not btn_xml or not btn_xml.is_visible():
            raise Exception(
                "ERROR_UBICACION_BOTON: No se pudo localizar el botón XML."
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
                interval_ms=150,
            )

            if not ready:
                raise Exception("TIMEOUT_DESCARGA_XML_SUNAT")

            if download_detect:
                download = download_detect[0]
                logger.info(f"[{self.area}] Archivo .zip descargado con éxito de SUNAT: {download.suggested_filename}")
                return download

            if modal_error_frame.is_visible() or modal_error_page.is_visible():
                logger.error(f"[{self.area}] Detectado 'Error del Servidor' de SUNAT al descargar XML. Abortando de inmediato...")
                raise Exception("SUNAT_ERROR_DEL_SERVIDOR")
        finally:
            page.remove_listener("download", capture_download)

    def Extract_XML_File(self, zip_path: Path) -> list[str]:
        extracted_xmls = []
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            for file_info in zip_ref.infolist():
                filename = Path(file_info.filename).name
                if filename.lower().endswith(".xml") and not filename.startswith("R-"):
                    destino = self.path_folder_xmls / filename
                    with (
                        zip_ref.open(file_info) as source,
                        open(destino, "wb") as target,
                    ):
                        shutil.copyfileobj(source, target)
                    extracted_xmls.append(filename)
        return extracted_xmls

    def Get_Dynamic_Quota(self) -> int:
        for old_quota in self.path_area_local.glob("Cuota_Diaria_*.json"):
            if old_quota.name != self.path_quota_file.name:
                old_quota.unlink(missing_ok=True)

        download_today = 0
        if self.path_quota_file.exists():
            try:
                data = json.loads(
                    self.path_quota_file.read_text(encoding="utf-8")
                )
                download_today = int(data.get("descargados_hoy", 0))
            except Exception as e:
                logger.warning(
                    f"[{self.area}] No se pudo leer {self.path_quota_file.name}, iniciando en 0: {e}"
                )

        today_hour = get_current_datetime().hour
        if today_hour < 12:
            time_slot = "MAÑANA (< 12:00)"
            turn_quota = QUOTA_MORNING
        elif today_hour < 18:
            time_slot = "TARDE (12:00 - 17:59)"
            turn_quota = QUOTA_AFTERNOON
        else:
            time_slot = "NOCHE (>= 18:00)"
            turn_quota = QUOTA_NIGHT

        remanente_diario = max(0, MAX_DAILY_QUOTA - download_today)
        effective_quota = min(turn_quota, remanente_diario)

        logger.info(
            f"[{self.area}] Franja {time_slot} | Cuota turno: {turn_quota} | "
            f"Descargados hoy: {download_today}/{MAX_DAILY_QUOTA} | "
            f"Cuota asignada a esta ejecución: {effective_quota}"
        )
        return effective_quota

    def Update_Daily_Quota(self, new_downloads: int) -> None:
        if new_downloads <= 0:
            return

        download_today = 0
        if self.path_quota_file.exists():
            try:
                data = json.loads(
                    self.path_quota_file.read_text(encoding="utf-8")
                )
                download_today = int(data.get("descargados_hoy", 0))
            except Exception:
                download_today = 0

        updated_total = download_today + new_downloads
        payload = {
            "fecha": self.period_time,
            "area": self.area,
            "descargados_hoy": updated_total,
            "ultima_actualizacion": get_current_datetime().isoformat(),
        }
        self.path_quota_file.write_text(
            json.dumps(payload, indent=2), encoding="utf-8"
        )
        logger.info(
            f"[{self.area}] Cuota diaria actualizada: {updated_total}/{MAX_DAILY_QUOTA} XMLs descargados hoy."
        )

    def Zip_Downloaded_XMLs(self, timestamp: str) -> list[tuple[Path, list[str]]]:
        files_xml = sorted(self.path_folder_xmls.glob("*.xml"))
        if not files_xml:
            raise Exception(
                f"[{self.area}] No se encontraron archivos XML para empaquetar en el ZIP de correo."
            )

        max_bytes = int(MAX_ZIP_SIZE_MB * 1024 * 1024)
        grupos_xml: list[list[Path]] = []
        grupo_actual: list[Path] = []

        buffer = io.BytesIO()
        zf_sim = zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED)

        for xml_file in files_xml:
            zf_sim.write(xml_file, arcname=xml_file.name)
            if buffer.tell() > max_bytes and grupo_actual:
                zf_sim.close()
                grupos_xml.append(grupo_actual)

                grupo_actual = [xml_file]
                buffer = io.BytesIO()
                zf_sim = zipfile.ZipFile(
                    buffer, "w", compression=zipfile.ZIP_DEFLATED
                )
                zf_sim.write(xml_file, arcname=xml_file.name)
            else:
                grupo_actual.append(xml_file)

        zf_sim.close()
        if grupo_actual:
            grupos_xml.append(grupo_actual)

        total_partes = len(grupos_xml)
        zips_generados: list[tuple[Path, list[str]]] = []

        for idx_parte, lista_archivos in enumerate(grupos_xml, start=1):
            if total_partes == 1:
                name_zip = f"XMLs_SUNAT_{self.area}_{timestamp}.zip"
            else:
                name_zip = f"XMLs_SUNAT_{self.area}_{timestamp}_PARTE_{idx_parte}_DE_{total_partes}.zip"

            path_zip_local = self.path_folder_xmls_zip / name_zip
            nombres_incluidos: list[str] = []

            with zipfile.ZipFile(path_zip_local, "w", compression=zipfile.ZIP_DEFLATED) as zip_out:
                for xml_file in lista_archivos:
                    zip_out.write(xml_file, arcname=xml_file.name)
                    nombres_incluidos.append(xml_file.name)

            size_mb = path_zip_local.stat().st_size / (1024 * 1024)
            logger.info(
                f"[{self.area}] ZIP generado ({idx_parte}/{total_partes}): {name_zip} "
                f"({size_mb:.2f} MB con {len(nombres_incluidos)} comprobantes)."
            )
            zips_generados.append((path_zip_local, nombres_incluidos))

        return zips_generados

    def Execute_Robot(self):
        global MAX_RETRIES_PER_RECORD

        self.file_log_exists = start_log(self.path_logs, self.file_log_exists, False)

        if not self.Acquire_Lock():
            logger.warning(f"[{self.area}] Ya existe una ejecución en curso. Omitiendo...")
            remove_log()
            return

        try:
            self.start_time = get_current_datetime()
            name_process = f"Descarga XMLs SUNAT - Área: {self.area}"
            self.console.rule(f"[grey66]Proceso Iniciado: [bold white]{name_process}[/bold white][/grey66]")

            for folder in [
                self.path_folder_inputs,
                self.path_folder_logs,
                self.path_folder_processed,
                self.path_folder_quarantine,
                self.path_folder_xmls,
            ]:
                folder.mkdir(parents=True, exist_ok=True)

            self.Ingest_Cloud_Inputs()

            logger.info("Iniciando Robot...")

            control_storage_day(
                [
                    self.path_folder_history_cloud,
                    self.path_area_cloud / "Xmls",
                    self.path_folder_xmls_zip_cloud,
                    self.path_folder_processed_cloud,
                    self.path_folder_quarantine_cloud,
                ],
                path_file_control=self.path_onedrive_lock,
            )

            remove_log()
            start_log(self.path_logs, self.file_log_exists, True)

            df_validos, df_observados, archivos_origen = self.Load_Clean_Data()

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_observados.empty:
                ruta_obs = (
                    self.path_folder_quarantine
                    / f"Comprobantes_Observados_{self.area}_{timestamp_actual}.xlsx"
                )
                export_df_to_excel(ruta_obs, df_observados[COLUMNS_FINAL])

                copy_items_to_folder(self.path_folder_quarantine,self.path_folder_quarantine_cloud,True,False,self.area)
                
                logger.warning(
                    f"[{self.area}] Se guardó Excel con registros observados en OneDrive: {self.path_folder_quarantine_cloud.name}/{ruta_obs.name} ({len(df_observados)} filas)"
                )

            if df_validos.empty:
                raise Exception(
                    f"[{self.area}] No se encontraron registros válidos para procesar."
                )

            total_rows = len(df_validos)
            execution_quota = self.Get_Dynamic_Quota()

            df_lote_actual = df_validos.iloc[:execution_quota].copy()

            df_remanentes_cola = df_validos.iloc[execution_quota:].copy()
            df_remanentes_cola["DETALLE_ESTADO"] = "NO PROCESADO - REMANENTE"
            df_remanentes_cola["FECHA_PROCESO"] = None

            total_to_process = len(df_lote_actual)
            logger.info(
                f"[{self.area}] Total disponible: {total_rows} registros. Cuota asignada para este ciclo: {total_to_process} registros."
            )

            map_xmls_locals = self.Index_XMLs_Locals()
            xml_por_idx: dict[int, str] = {}
            new_sunat_downloads = 0

            contador_batch_bk = 0
            for idx, fila in df_lote_actual.iterrows():
                contador_batch_bk += 1
                ruc = str(fila["RUC"]).strip()
                serie = str(fila["SERIEDOC"]).strip()
                numero = str(fila["NUMERODOC"]).strip()
                num_limpio = str(int(numero)) if numero.isdigit() else numero

                identificador_cpe = f"{serie}-{numero} (RUC: {ruc})"

                file_exists = map_xmls_locals.get((ruc, serie, num_limpio)) or map_xmls_locals.get(("", serie, num_limpio))
                if file_exists:
                    logger.info(
                        f"[{contador_batch_bk}/{total_rows}] OMITIENDO: {identificador_cpe} ya existe en disco ({file_exists})."
                    )
                    df_lote_actual.at[idx, "ESTADO"] = "DESCARGADO"
                    df_lote_actual.at[idx, "DETALLE_ESTADO"] = (
                        f"RECUPERADO_LOCAL_PREVIO: {file_exists}"
                    )
                    df_lote_actual.at[idx, "FECHA_PROCESO"] = pd.Timestamp.now()
                    xml_por_idx[idx] = file_exists

            df_pendientes_sunat = df_lote_actual[
                df_lote_actual["ESTADO"] == "PENDIENTE"
            ]

            contador_batch = 0

            if not df_pendientes_sunat.empty:
                with sync_playwright() as p:
                    for current_company, df_grupo in df_pendientes_sunat.groupby("EMPRESA", sort=False):
                        if current_company == "LPE":
                            MAX_RETRIES_PER_RECORD = 4
                        else:
                            MAX_RETRIES_PER_RECORD = 3

                        logger.info(
                            f"[{self.area}] Iniciando bloque para empresa {current_company} ({len(df_grupo)} comprobantes)..."
                        )
                        context, page, frame = None, None, None
                        sunat_colapse = False

                        try:
                            context, page, frame = self.Login_Sunat(p, current_company)
                        except SunatConexionError as sce:
                            logger.error(
                                f"[{self.area} | {current_company}] No se pudo iniciar sesión:\n{sce}"
                            )
                            continue

                        for idx, fila in df_grupo.iterrows():
                            contador_batch += 1
                            filtro = (
                                str(fila.get("FILTRO", "Recibido")).strip().capitalize()
                                or "Recibido"
                            )
                            ruc = str(fila["RUC"]).strip()
                            tipo = str(fila["COD_DCMNTO"]).strip()
                            serie = str(fila["SERIEDOC"]).strip()
                            numero = str(fila["NUMERODOC"]).strip()

                            identificador_cpe = (
                                f"[{self.area}|{current_company}] {serie}-{numero} (RUC: {ruc})"
                            )
                            logger.info(
                                f"[{contador_batch}/{len(df_pendientes_sunat)}] Consultando SUNAT: {identificador_cpe}"
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
                                        r"servicios de SUNAT|reintentar en 5 minutos", re.IGNORECASE,
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
                                        interval_ms=150,
                                    )

                                    if not respuesta_servidor:
                                        raise Exception("SUNAT_SPINNER_TIMEOUT_BLOQUEADO")
                                    

                                    if modal_error_frame.is_visible() or modal_error_page.is_visible() or modal_error_page_aux.is_visible():
                                        logger.error(
                                            f"[{self.area}|{current_company}] Detectado 'Error del Servidor' de SUNAT en intento {intento}..."
                                        )
                                        raise Exception("SUNAT_ERROR_DEL_SERVIDOR")

                                    if alerta_cpe_no_existe.is_visible():
                                        texto_alerta = (alerta_cpe_no_existe.inner_text())
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
                                        texto_alerta = (alerta_mayor_24.inner_text())
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

                                        detalle_resultado = f"HISTORICO_24_MESES_PEDIDO_DESCARGA_NRO_{ticket_id}"
                                        logger.warning(f"[{self.area}|{current_company}] Comprobante > 24 meses. N° Pedido: {ticket_id}")
                                        
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
                                    temp_zip_path = self.path_folder_xmls / download.suggested_filename
                                    download.save_as(temp_zip_path)

                                    archivos_ok = self.Extract_XML_File(temp_zip_path)
                                    temp_zip_path.unlink(missing_ok=True)

                                    if archivos_ok:
                                        descarga_exitosa = True
                                        new_sunat_downloads += 1
                                        xml_por_idx[idx] = archivos_ok[0]
                                        detalle_resultado = (f"DESCARGA_EXITOSA: {archivos_ok[0]}")
                                        logger.info(f"[{self.area}|{current_company}] Descarga y extracción XML exitosa: {archivos_ok[0]}")
                                    else:
                                        detalle_resultado = "ZIP_SIN_XML_VALIDO"

                                    frame.get_by_role("button", name="Close").click()
                                    frame.get_by_role("button", name="Limpiar").click()
                                    break
                                except Exception as ex:
                                    logger.warning(
                                        f"[{self.area}|{current_company}] Fallo en intento {intento}/{MAX_RETRIES_PER_RECORD} para {identificador_cpe}: {ex}"
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
                                            logger.info("Esperando unos segundos para que los servicios de SUNAT respondan...")
                                            time.sleep(1.5)
                                            context, page, frame = self.Login_Sunat(p, current_company)
                                        else:
                                            context, page, frame = self.Recover_Screen_Restart(p, context, page, current_company)
                                    except SunatConexionError as sce:
                                        logger.error(
                                            f"[{self.area} | {current_company}] SUNAT colapsó. Deteniendo bloque: {sce}"
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

                            get_random_delay(BREAK_MIN_CALL,BREAK_MAX_CALL)

                            if (contador_batch % BATCH_SIZE == 0) and (contador_batch < len(df_pendientes_sunat)):
                                logger.info(
                                    f"[{self.area}] Lote de {BATCH_SIZE} alcanzado. Pausando {BREAK_MINUTE_BATCH} min..."
                                )
                                time.sleep(BREAK_MINUTE_BATCH * 60)

                        self.Close_Sunat_Session(context, page, current_company)

            self.Update_Daily_Quota(new_sunat_downloads)

            df_descargados = df_lote_actual[df_lote_actual["ESTADO"] == "DESCARGADO"].copy()

            observados_sunat = df_lote_actual["DETALLE_ESTADO"].str.contains(
                r"COMPROBANTE_NO_EXISTE|HISTORICO_24_MESES", regex=True, na=False,
            )
            df_observados_sunat = df_lote_actual[observados_sunat].copy()
            if not df_observados_sunat.empty:
                df_observados_sunat["ESTADO"] = "OBSERVADO_SUNAT"

            df_fallidos_tecnicos = df_lote_actual[
                (
                    (df_lote_actual["ESTADO"] == "NO_DESCARGADO")
                    | (df_lote_actual["ESTADO"] == "PENDIENTE")
                )
                & (~observados_sunat)
            ].copy()

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_observados_sunat.empty:
                if not self.path_folder_quarantine.exists():
                    self.path_folder_quarantine.mkdir(parents=True, exist_ok=True)
                ruta_obs_sunat = (
                    self.path_folder_quarantine
                    / f"Comprobantes_No_Descargados_SUNAT_{self.area}_{timestamp_actual}.xlsx"
                )
                export_df_to_excel(ruta_obs_sunat, df_observados_sunat[COLUMNS_FINAL])
                copy_items_to_folder(self.path_folder_quarantine,self.path_folder_quarantine_cloud,True,False,self.area)
                logger.warning(
                    f"[{self.area}] Se guardó Excel con registros observados por SUNAT en OneDrive: {self.path_folder_quarantine_cloud.parent}/{self.path_folder_quarantine_cloud.name} ({len(df_observados_sunat)} filas)"
                )

            for archivo_leido in archivos_origen:
                if archivo_leido.exists():
                    archivo_leido.unlink()
                    logger.info(f"[{self.area}] Se eliminó archivo procesado: {archivo_leido.name}")

            df_nueva_cola = pd.concat(
                [df_fallidos_tecnicos, df_remanentes_cola], ignore_index=True
            )

            timestamp_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not df_nueva_cola.empty:
                ruta_cola_inputs = self.path_folder_inputs / f"Comprobantes_Pendientes_{self.area}_{timestamp_actual}.xlsx"
                df_nueva_cola["ESTADO"] = "PENDIENTE"
                export_df_to_excel(ruta_cola_inputs, df_nueva_cola[COLUMNS_FINAL])
                logger.warning(
                    f"[{self.area}] Quedan {len(df_nueva_cola)} comprobantes pendientes ({len(df_fallidos_tecnicos)} por reintentar + {len(df_remanentes_cola)} fuera de cuota)."
                    f"[{self.area}] Actualizado en: {ruta_cola_inputs.name} para la próxima ejecución."
                )
            else:
                logger.info(f"[{self.area}] Se descargó los XMLs de todos los comprobantes. Carpeta Inputs limpia...")

            email_failed = True
            emails_sent_count = 0

            if not df_descargados.empty:
                copy_items_to_folder(self.path_folder_xmls,self.path_folder_xmls_cloud,False,True,self.area)

                logger.info(
                    f"[{self.area}] Se guardó Lote de XMLs descargados a OneDrive: {self.path_folder_xmls_cloud.parent}/{self.path_folder_xmls_cloud.name} ({len(df_descargados)} archivos)"
                )

                self.path_folder_xmls_zip.mkdir(parents=True, exist_ok=True)
                zips_generados = self.Zip_Downloaded_XMLs(timestamp_actual)
                total_parts = len(zips_generados)

                if total_parts > 0:
                    for idx_parte, (ruta_zip, nombres_xml_parte) in enumerate(zips_generados, start=1):
                        set_xmls_parte = set(nombres_xml_parte)
                        idxs_en_parte = [
                            i for i, xml_name in xml_por_idx.items()
                            if xml_name in set_xmls_parte and i in df_descargados.index
                        ]

                        part_companies = sorted(
                            df_descargados.loc[idxs_en_parte, "EMPRESA"]
                            .unique()
                            .tolist()
                        )
                    
                        if total_parts == 1:
                            email_subject = f"ROBOT SUNAT XML | [{self.area}] | Lote {self.period_time} ({len(nombres_xml_parte)} XMLs)"
                        else:
                            email_subject = (
                                f"ROBOT SUNAT XML | {self.area} | Lote {self.period_time} — "
                                f"Parte {idx_parte} de {total_parts}"
                            )

                        body_html = build_sunat_email_html(
                            self.area,
                            self.period_time,
                            len(nombres_xml_parte),
                            len(df_descargados),
                            ruta_zip.name,
                            part_companies,
                            idx_parte,
                            total_parts,
                        )

                        email_sent = send_email_graph(
                            TENANT_ID,
                            CLIENT_ID,
                            CLIENT_SECRET,
                            MAIL_BOX,
                            RECIPIENTS_MAIL_LIST,
                            email_subject,
                            body_html,
                            ruta_zip,
                            self.area,
                        )

                        if email_sent:
                            email_failed = False
                            emails_sent_count += 1
                            df_descargados.loc[idxs_en_parte, "ESTADO"] = "ENVIADO_CORREO"
                            df_descargados.loc[idxs_en_parte, "DETALLE_ESTADO"] = df_descargados.loc[idxs_en_parte, "DETALLE_ESTADO"] + f" | ENVIADO_ZIP: {ruta_zip.name}"
                        else:
                            df_descargados.loc[idxs_en_parte, "DETALLE_ESTADO"] = df_descargados.loc[idxs_en_parte, "DETALLE_ESTADO"] + f" | ERROR_ENVIO_ZIP: {ruta_zip.name}"
                            logger.error(
                                f"[{self.area}] No se pudo enviar el correo para el paquete {ruta_zip.name}."
                            )

                    copy_items_to_folder(self.path_folder_xmls_zip, self.path_folder_xmls_zip_cloud,True,False,self.area)
                    logger.info(
                        f"[{self.area}] Se guardaron {total_parts} paquete(s) ZIP BK en OneDrive: {self.path_folder_xmls_zip_cloud}"
                    )

                    if self.path_folder_xmls_zip.exists():
                        remove_temporal_folder(self.path_folder_xmls_zip,self.area)

                path_process = self.path_folder_processed / f"Comprobantes_Descargados_{self.area}_{timestamp_actual}.xlsx"
                export_df_to_excel(path_process, df_descargados[COLUMNS_FINAL])

                copy_items_to_folder(self.path_folder_processed,self.path_folder_processed_cloud,True,False,self.area)
                logger.info(
                    f"[{self.area}] Se guardó Excel con registros procesados en OneDrive: {self.path_folder_processed_cloud.parent}/{self.path_folder_processed_cloud.name} ({len(df_descargados)} filas)"
                )

                if self.path_folder_xmls.exists():
                    remove_temporal_folder(self.path_folder_xmls,self.area)
            else:
                email_failed = False
                logger.warning(f"No se envío correo al no encontrar XMLs descargados.")

            if self.path_folder_xmls.exists():
                if list(self.path_folder_xmls.iterdir()):
                    remove_temporal_folder(self.path_folder_xmls,self.area)

            if self.path_folder_xmls_zip.exists():
                if list(self.path_folder_xmls_zip.iterdir()):
                    remove_temporal_folder(self.path_folder_xmls_zip,self.area)

            if email_failed and not df_descargados.empty:
                self.Finish_Process(
                    False,
                    f"Falló el envío de {"algunos correos" if emails_sent_count > 1 else "correo"}.",
                )
            else:
                self.Finish_Process(True)
        except Exception as e:
            self.Finish_Process(False, str(e))
        finally:
            self.Release_Lock()

    def Finish_Process(self, is_success: bool = False, msg_error: str | None = None):
        self.end_time = get_current_datetime()

        if is_success:
            logger.success(f"[{self.area}] Ejecución exitosa: Se procesó la información.")
            close_rule = lambda: self.console.rule(
                f"[grey66]Proceso Finalizado ({self.area})[/grey66]"
            )
        else:
            logger.error(f"[{self.area}] Proceso Incompleto. Detalle: {msg_error}")
            close_rule = lambda: show_custom_rule(self.console, f"Proceso Finalizado con Error ({self.area})", state="Error")

        total_seconds = int((self.end_time - self.start_time).total_seconds())
        minutes, seconds = divmod(total_seconds, 60)
        difference_formatted = f"{minutes} minuto(s), {seconds} segundo(s)"

        remove_log()
        self.file_log_exists = add_log_file(self.path_logs, self.file_log_exists, True)
        logger.info(f"[{self.area}] Tiempo de proceso: {difference_formatted}")
        add_log_console()
        print(f"[dark_orange][{self.area}] Tiempo de proceso: {difference_formatted}[/dark_orange]")
        
        close_rule()

def run_worker_area(area_nombre: str):
    Robot_Sunat(area_nombre)

if __name__ == "__main__":
    multiprocessing.freeze_support()

    if len(sys.argv) > 1:
        area_arg = sys.argv[1].strip().capitalize()
        if area_arg in AVAILABLE_AREAS:
            run_worker_area(area_arg)
        else:
            print(
                f"[bold red]Área '{area_arg}' no válida. Opciones: {AVAILABLE_AREAS}[/bold red]"
            )
    else:
        procesos: list[multiprocessing.Process] = []
        for area in AVAILABLE_AREAS:
            p = multiprocessing.Process(
                target=run_worker_area, args=(area,), name=f"Worker_{area}"
            )
            p.start()
            procesos.append(p)
            time.sleep(5)

        for p in procesos:
            p.join()