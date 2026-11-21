import os
import stat
import subprocess
import time
import ctypes
from datetime import datetime, date, timedelta
import sys
import shutil
from pathlib import Path
from zoneinfo import ZoneInfo

from loguru import logger
from rich import print
from rich.text import Text
import pandas as pd
import msal
import requests
import base64

IS_DEV = os.getenv("APP_ENV", "PROD") == "DEV"

LIMA_TZ = ZoneInfo("America/Lima")

def get_current_datetime():
    return datetime.now(tz=LIMA_TZ)

def show_custom_rule(console, titulo, state = "Success"):
    total_width = console.width
    if state == "Success":
        color_line = "bold green"
        color_text = "grey66"
    elif state == "Error":
        color_line = "red"
        color_text = "grey66"
    else:
        color_line = "cyan"
        color_text = "grey66"

    text = f" {titulo} "
    length_text = len(Text.from_markup(text).plain)

    length_line = max((total_width - length_text) // 2, 0)
    line = "─" * length_line

    rule = f"[{color_line}]{line}[/{color_line}][{color_text}]{text}[/{color_text}][{color_line}]{line}[/{color_line}]"
    console.print(rule)
    
def custom_format(type_process: int):
    def formatter(record: dict):
        levelname = record["level"].name
        if levelname == "INFO":
            text = "AVISO"
            level_str = f"<cyan>{text:<7}</cyan>"
            message_color = "<cyan>"
        elif levelname == "WARNING":
            text = "ALERTA"
            level_str = f"<level>{text:<7}</level>"
            message_color = "<level>"
        elif levelname == "SUCCESS":
            text = "ÉXITO"
            level_str = f"<level>{text:<7}</level>"
            message_color = "<level>"
        else:
            level_str = f"<level>{levelname:<7}</level>"
            message_color = "<level>"
        
        original_message = str(record["message"])
        safe_message = original_message.replace("{", "{{").replace("}", "}}")
        custom_message = f"{message_color}{safe_message}</{message_color.strip('<>')}>\n"
        
        if type_process == 0:
            level_str = f"{level_str} | "
        else:
            level_str = f"{level_str} | {record['name']}:{record['function']}:{record['line']} - "
            if record["exception"] is not None:
                custom_message += f"{record['exception']}\n"

        return (
            f"<cyan><bold>{record['time']:DD/MM/YYYY HH:mm:ss}</bold></cyan> | "
            f"{level_str}"
            f"{custom_message}"
        )
    return formatter

def remove_log():
    logger.remove()

def add_log_console():
    logger.add(sys.stdout,
            backtrace=False, diagnose=IS_DEV, level="DEBUG",
            colorize=True,
            format=custom_format(0))

def add_log_file(path_log: Path, file_log_exist: bool, exits_log: bool):
    if path_log.exists() and not exits_log:
        logger.add(path_log, 
                backtrace=True, diagnose=IS_DEV, level="INFO",
                format="\n\n{time:DD/MM/YYYY HH:mm:ss} | {level:<7} | {name}:{function}:{line} - {message}") 
        return file_log_exist
    
    logger.add(path_log, 
        backtrace=True, diagnose=IS_DEV, level="INFO",
        format="{time:DD/MM/YYYY HH:mm:ss} | {level:<7} | {name}:{function}:{line} - {message}") 
    return True

def start_log(path_log: Path, file_log_exist: bool, exits_log: bool = False):
    remove_log()
    add_log_console()
    file_log_exist = add_log_file(path_log, file_log_exist, exits_log)
    return file_log_exist

def export_df_to_excel(path_file: Path, df: pd.DataFrame):    
    def format_columns(df,worksheet):
        for idx, col in enumerate(df):  
            series = df.iloc[:, idx]

            header_str = str(series.name) if series.name is not None else ""
            header_len = len(header_str)

            max_val_len = series.astype(str).str.len().max()
            max_val_len = 0 if pd.isna(max_val_len) else int(max_val_len)

            max_len = max(max_val_len, header_len) + 3
            worksheet.set_column(idx, idx, max_len)
    
    writer = pd.ExcelWriter(path_file, engine="xlsxwriter", date_format="dd/mm/yyyy", datetime_format="dd/mm/yyyy hh:mm:ss")

    df.to_excel(writer, index=False, sheet_name="Reporte")
    
    worksheet = writer.sheets["Reporte"]
    
    (max_row, max_col) = df.shape
    column_settings = [{"header": column} for column in df.columns]

    if max_row>0:
        worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings,
                                                            "style": "Table Style Light 14"})
        format_columns(df,worksheet)

    writer.close()

def handle_remove_readonly(func, path: str, exc: BaseException) -> None:
    if not isinstance(exc, PermissionError) or func not in (os.rmdir, os.unlink):
        raise exc

    os.chmod(path, stat.S_IWRITE)
    for attempt in range(3):
        try:
            func(path)
            return
        except PermissionError:
            if attempt == 2:
                raise
            time.sleep(0.2)

def remove_temporal_folder(path_folder: Path, area: str = "Siniestros"):
    try:
        shutil.rmtree(
            path_folder,
            ignore_errors=False,
            onexc=handle_remove_readonly,
        )
    except Exception as e:
        logger.warning(
            f"[{area}] No se pudo eliminar la carpeta temporal {path_folder}: {e}"
        )

def copy_items_to_folder(path_folder_src: Path, path_folder_dtn: Path, clear_folder_src : bool = False, create_diff_folder_dtn : bool = False, area: str = "Siniestros"):    
    path_folder_dtn_new: Path = path_folder_dtn
    
    if not path_folder_dtn.exists():
        path_folder_dtn.mkdir(parents=True, exist_ok=True)
    else:
        if create_diff_folder_dtn:
            # hour_seconds = get_current_datetime()
            # path_folder_dtn = Path(str(path_folder_dtn) + "_" + hour_seconds.strftime("%H%M%S"))
            path_folder_dtn_new = path_folder_dtn.parent /  f"{path_folder_dtn.name}_2"
            if path_folder_dtn_new.exists():
                path_folder_list = [
                    int(str((subfolder.name).split("_")[2]).strip()) for subfolder in path_folder_dtn.parent.iterdir() 
                    if subfolder.is_dir() and path_folder_dtn.name in subfolder.name 
                    and len((subfolder.name).split("_")) > 2
                    and str((subfolder.name).split("_")[2]).strip().isdigit()
                ]
                path_folder_list.sort(reverse=True)
                suffix = path_folder_list[0] + 1
                path_folder_dtn_new = path_folder_dtn.parent /  f"{path_folder_dtn.name}_{suffix}"
            path_folder_dtn_new.mkdir(parents=True, exist_ok=True)

    for item in path_folder_src.iterdir():
        dst_item = path_folder_dtn_new / item.name

        if not dst_item.exists():
            shutil.copy2(item, dst_item)

    if clear_folder_src:
        remove_temporal_folder(path_folder_src, area)

def build_sunat_email_html(
    area: str,
    process_date: str,
    xmls_in_package: int,
    total_lote: int,
    zip_name: str,
    included_companies: list[str],
    current_part: int = 1,
    total_parts: int = 1,
) -> str:
    empresas_str = " / ".join(included_companies) if included_companies else "-"
    badge_entrega = (
        f"PAQUETE {current_part} DE {total_parts}"
        if total_parts > 1
        else "ENTREGA ÚNICA"
    )

    fila_consolidado = ""
    if total_parts > 1:
        fila_consolidado = f"""
        <tr>
            <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #4B5563; font-weight: 600; width: 42%;">Total consolidado del lote</td>
            <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #111827;">{total_lote} comprobantes</td>
        </tr>
        """

    return f"""
    <!DOCTYPE html>
    <html lang="es">
    <head><meta charset="UTF-8"></head>
    <body style="margin: 0; padding: 0; background-color: #F3F4F6; font-family: 'Segoe UI', Arial, sans-serif;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color: #F3F4F6; padding: 24px 12px;">
            <tr>
                <td align="center">
                    <table role="presentation" width="600" cellspacing="0" cellpadding="0" style="background-color: #FFFFFF; border-radius: 8px; overflow: hidden; border: 1px solid #E5E7EB;">
                        <tr>
                            <td style="background-color: #ff6d46; padding: 20px 24px;">
                                <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td>
                                            <span style="display: inline-block; background-color: #029790; color: #E0F2FE; font-size: 11px; font-weight: 700; letter-spacing: 0.8px; padding: 4px 10px; border-radius: 4px; text-transform: uppercase;">
                                                ROBOT_SUNAT_PYTHON
                                            </span>
                                            <h1 style="margin: 10px 0 2px 0; color: #FFFFFF; font-size: 19px; font-weight: 600;">
                                                Descarga Automatizada de XMLs SUNAT
                                            </h1>
                                        </td>
                                        <td align="right" valign="top">
                                            <span style="display: inline-block; background-color: #7878be; color: #FFFFFF; font-size: 11px; font-weight: 700; padding: 5px 10px; border-radius: 4px;">
                                                {badge_entrega}
                                            </span>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>

                        <tr>
                            <td style="padding: 20px 24px 12px 24px; color: #374151; font-size: 14px; line-height: 1.5;">
                                Estimados,<br>
                                Se adjunta el archivo comprimido con los comprobantes electrónicos (XML) descargados desde el portal SUNAT SOL correspondientes al proceso de <strong>{area}</strong>.
                            </td>
                        </tr>

                        <tr>
                            <td style="padding: 8px 24px 0px 24px;">
                                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #E5E7EB; border-radius: 6px; border-collapse: collapse; font-size: 13px;">
                                    <tr style="background-color: #F8FAFC;">
                                        <th colspan="2" align="left" style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #7878be; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">
                                            Resumen del Paquete Adjunto
                                        </th>
                                    </tr>
                                    <tr>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #4B5563; font-weight: bold; width: 42%;">🗄️ Área</td>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #111827;">{area}</td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #4B5563; font-weight: bold;">🏢 Empresa(s)</td>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #111827;">{empresas_str}</td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #4B5563; font-weight: bold;">🗓️ Fecha de Proceso</td>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #111827;">{process_date}</td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #4B5563; font-weight: bold;">📋 XMLs en este archivo</td>
                                        <td style="padding: 10px 14px; border-bottom: 1px solid #E5E7EB; color: #059669; font-weight: bold;">{xmls_in_package} comprobantes</td>
                                    </tr>
                                    {fila_consolidado}
                                    <tr>
                                        <td style="padding: 10px 14px; color: #4B5563; font-weight: bold;">📂 Archivo Adjunto</td>
                                        <td style="padding: 10px 14px; color: #1F2937; font-family: Consolas, monospace; font-size: 12px;">{zip_name}</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>

                        <tr>
                            <td style="background-color: #fcfaf8; padding: 14px 5px; text-align: center;">
                                <span style="font-size: 12px; color: #9ca3af;">
                                Este es un mensaje generado automáticamente <strong>ROBOT_SUNAT_PYTHON</strong>. <br>Por favor, no responder a este correo. Equipo Automatización & IA.
                                </span>
                            </td>
                        </tr>

                        <tr>
                            <td>
                                <table style="width:100%; margin:0; background-color:#ff6d46; padding:10px 25px">
                                    <tbody>
                                        <tr>  <!-- valign="bottom" -->
                                            <td style="width:50%; padding:0">
                                                <table>
                                                    <tbody>
                                                        <tr>
                                                            <td><a href="https://www.facebook.com/LaPositivaSeguros/"
                                                                    target="_blank" rel="noopener noreferrer"
                                                                    data-auth="NotApplicable" style="text-decoration:none"
                                                                    data-linkindex="0"
                                                                    title="https://www.facebook.com/LaPositivaSeguros/"><img
                                                                        data-imagetype="External"
                                                                        src="http://cdn.email-platform.com/positivaseguros/COBRANZAS/Estaticas/facebook.png"
                                                                        data-imageproxyendpoint="/actions/ei" alt="facebook"
                                                                        style="max-width:10px; padding-right:5px"> </a><a
                                                                    href="https://twitter.com/lapositiva_peru?lang=es"
                                                                    target="_blank" rel="noopener noreferrer"
                                                                    data-auth="NotApplicable" style="text-decoration:none"
                                                                    data-linkindex="1"
                                                                    title="https://twitter.com/lapositiva_peru?lang=es"><img
                                                                        data-imagetype="External"
                                                                        src="https://images.email-platform.com/positivaseguros/COBRANZAS/Estaticas/twitter_small.png"
                                                                        alt="twitter" style="max-width:10px; padding-right:5px">
                                                                </a><a href="https://www.youtube.com/user/LaPositivaPeru"
                                                                    target="_blank" rel="noopener noreferrer"
                                                                    data-auth="NotApplicable" style="text-decoration:none"
                                                                    data-linkindex="2"
                                                                    title="https://www.youtube.com/user/LaPositivaPeru"><img
                                                                        data-imagetype="External"
                                                                        src="http://cdn.email-platform.com/positivaseguros/COBRANZAS/Estaticas/youtube.png"
                                                                        data-imageproxyendpoint="/actions/ei" alt="youtube"
                                                                        style="max-width:10px; padding-right:5px"> </a></td>
                                                        </tr>
                                                        <tr>
                                                            <td
                                                                style="font-family:Helvetica; font-size:13px; font-weight:bolder">
                                                                <a href="https://www.lapositiva.com.pe/wps/portal/corporativo/home"
                                                                    target="_blank" rel="noopener noreferrer"
                                                                    data-auth="NotApplicable"
                                                                    style="color:white!important; text-decoration:none!important"
                                                                    data-linkindex="3"
                                                                    title="https://www.lapositiva.com.pe/wps/portal/corporativo/home">lapositiva.pe</a>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="width:50%; padding:0; text-align:right">
                                                <a
                                                    href="https://www.lapositiva.com.pe/" target="_blank"
                                                    rel="noopener noreferrer" data-auth="NotApplicable" data-linkindex="4"
                                                    title="https://www.lapositiva.com.pe/"><img data-imagetype="External"
                                                        src="http://cdn.email-platform.com/positivaseguros/COBRANZAS/Estaticas/logo.png"
                                                        data-imageproxyendpoint="/actions/ei" alt="La Positiva"
                                                        style="max-width:100px; width:100%"> 
                                                </a>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """

def send_email_graph(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    mailbox: str,
    recipients: list[str],
    subject: str,
    body_text: str,
    path_file_zip: Path | None = None,
    area: str | None = None,
    is_html: bool = True,
) -> bool:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret, timeout=15
    )

    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" not in result:
        logger.error(f"Fallo al autenticar en Entra ID: {result.get('error_description')}")
        return False

    token = result["access_token"]

    to_recipients = [
        {"emailAddress": {"address": correo.strip()}}
        for correo in recipients
    ]

    mensaje = {
        "subject": subject,
        "body": {
            "contentType": "HTML" if is_html else "Text", 
            "content": body_text
        },
        "toRecipients": to_recipients,
    }

    if path_file_zip and path_file_zip.exists():
        with open(path_file_zip, "rb") as f:
            contenido_b64 = base64.b64encode(f.read()).decode("utf-8")

        mensaje["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": path_file_zip.name,
                "contentType": "application/zip",
                "contentBytes": contenido_b64,
            }
        ]

    url_envio = f"https://graph.microsoft.com/v1.0/users/{mailbox}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    payload = {"message": mensaje, "saveToSentItems": "true"}

    response = requests.post(url_envio, headers=headers, json=payload, timeout=(10, 30))

    if response.status_code == 202:
        logger.info(f"[{area}] Correo enviado exitosamente a {recipients}.")
        return True
    else:
        logger.error(f"[{area}] Error al enviar correo ({response.status_code}): {response.text}")
        return False

def cleanup_storage_onedrive(path_folder: Path) -> None:
    if not path_folder.exists():
        print(f"[Aviso] La ruta no existe: {path_folder}")
        return

    patron = str(path_folder / "*")

    # +U : Marca como Unpinned (Solo en línea / Liberar espacio)
    # -P : Remueve Pinned (Desancla el archivo local)
    # /s : Procesa recursivamente todos los archivos y subcarpetas
    # /d : Procesa también directorios
    comando = ["attrib", "+U", "-P", patron, "/s", "/d"]

    try:
        resultado = subprocess.run(
            comando, capture_output=True, text=True, check=False
        )
        if resultado.returncode == 0:
            print(
                f"[OneDrive] Solicitud de liberación de espacio carpeta '{path_folder.name}' enviada correctamente."
            )
        else:
            print(f"[OneDrive] Mensaje del sistema: {resultado.stderr.strip()}")
    except Exception as e:
        print(f"[OneDrive] Error al ejecutar attrib: {e}")

def control_storage_day(list_path_folder_onedrive: list[Path], path_file_control: Path = "onedrive_cleanup.lock") -> None:
    today = date.today().isoformat()

    path_file_control = Path(path_file_control)

    if path_file_control.exists():
        last_date_clear = path_file_control.read_text(
            encoding="utf-8"
        ).strip()
        if last_date_clear == today:
            print(
                f"[OneDrive] Mantenimiento ya realizado hoy ({today}). Saltando paso."
            )
            return

    print(f"[OneDrive] Primera ejecución del día ({today}). Liberando espacio...")

    for path_folder in list_path_folder_onedrive:
        cleanup_storage_onedrive(path_folder)

    path_file_control.write_text(today, encoding="utf-8")

def is_pid_running(pid: int) -> bool:
    if pid <= 0:
        return False

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    STILL_ACTIVE = 259
    ERROR_ACCESS_DENIED = 5

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    handle = kernel32.OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION, False, pid
    )

    if not handle:
        return ctypes.get_last_error() == ERROR_ACCESS_DENIED

    try:
        exit_code = ctypes.c_ulong()
        if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            return exit_code.value == STILL_ACTIVE
        return True
    finally:
        kernel32.CloseHandle(handle)