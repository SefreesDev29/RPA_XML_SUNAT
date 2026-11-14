import subprocess
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
            backtrace=False, diagnose=False, level="DEBUG",
            colorize=True,
            format=custom_format(0))

def add_log_file(path_log: Path, file_log_exist: bool, exits_log: bool):
    if path_log.exists() and not exits_log:
        logger.add(path_log, 
                backtrace=True, diagnose=True, level="INFO",
                format="\n\n{time:DD/MM/YYYY HH:mm:ss} | {level:<7} | {name}:{function}:{line} - {message}") 
        return file_log_exist
    
    logger.add(path_log, 
        backtrace=True, diagnose=True, level="INFO",
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

def copy_items_to_folder(path_folder_src: Path, path_folder_dtn: Path, clear_folder_src : bool = False, create_diff_folder_dtn : bool = False):
    if not path_folder_dtn.exists():
        path_folder_dtn.mkdir(parents=True)
    else:
        if create_diff_folder_dtn:
            # hour_seconds = get_current_datetime()
            # path_folder_dtn = Path(str(path_folder_dtn) + "_" + hour_seconds.strftime("%H%M%S"))
            path_items_list = (path_folder_dtn.name).split("_")
            if len(path_items_list) == 2:
                path_folder_dtn = path_folder_dtn.parent /  f"{path_folder_dtn.name}_2"
            else:
                suffix = int(path_items_list[2]) + 1
                path_folder_dtn = path_folder_dtn.parent /  f"{path_items_list[1]}_{suffix}"
            path_folder_dtn.mkdir()

    for item in path_folder_src.iterdir():
        dst_item = path_folder_dtn / item.name

        if not dst_item.exists():
            shutil.copy2(item, dst_item)

    if clear_folder_src:
        shutil.rmtree(path_folder_src)

def send_email_graph(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    mailbox: str,
    recipients: list[str],
    subject: str,
    body_text: str,
    path_file_zip: Path | None = None,
) -> bool:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
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
        "body": {"contentType": "Text", "content": body_text},
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

    response = requests.post(url_envio, headers=headers, json=payload)

    if response.status_code == 202:
        logger.success(f"Correo enviado exitosamente a {recipients}.")
        return True
    else:
        logger.error(f"Error al enviar correo ({response.status_code}): {response.text}")
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