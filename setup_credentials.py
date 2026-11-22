# import getpass getpass.getpass("Ingrese el Usuario SOL LPG: ")
import os
import sys
import keyring
import msvcrt
import time
from rich.console import Console
from rich.panel import Panel

# uv run --with nuitka python -m nuitka --onefile --assume-yes-for-downloads --include-package=keyring.backends --output-filename=Setup_Credentials.exe setup_credentials.py

console = Console()

AMBIENTE = os.getenv("APP_ENV", "DEV").upper()
SERVICIO = f"RPA_SUNAT_{AMBIENTE}"

def masked_input(prompt: str) -> str:
    sys.stdout.write(prompt)
    sys.stdout.flush()

    chars = []
    while True:
        key = msvcrt.getwch()
        if key in ("\r", "\n"):
            sys.stdout.write("\n")
            sys.stdout.flush()
            break
        elif key == "\x03":
            raise KeyboardInterrupt
        elif key == "\x08":
            if chars:
                chars.pop()
                sys.stdout.write("\b \b")
                sys.stdout.flush()
        elif key in ("\x00", "\xe0"):
            msvcrt.getwch()
        else:
            chars.append(key)
            sys.stdout.write("*")
            sys.stdout.flush()

    return "".join(chars)

console.print(
    Panel.fit(
        f"[bold white]APROVISIONAMIENTO DE CREDENCIALES[/bold white]\n"
        f"Entorno destino: [bold cyan]{SERVICIO}[/bold cyan]\nLas credenciales se cifrarán localmente.",
        border_style="cyan",
    )
)

try:
    usuario_lpg = masked_input("Ingrese el Usuario SOL LPG: ").strip()
    password_lpg = masked_input("Ingrese la Clave SOL LPG: ").strip()

    usuario_lpe = masked_input("Ingrese el Usuario SOL LPE: ").strip()
    password_lpe = masked_input("Ingrese la Clave SOL LPE: ").strip()

    tenant_id = masked_input("Ingrese el Tenant ID: ").strip()
    client_id = masked_input("Ingrese la Client ID: ").strip()
    client_secret = masked_input("Ingrese la Client Secret (Value): ").strip()

    if (not usuario_lpg or not password_lpg or not usuario_lpe or not password_lpe 
        or not tenant_id or not client_id or not client_secret):
        console.print(
            "\n[bold red][ERROR] Operación cancelada: ningún campo puede estar vacío.[/bold red]"
        )
        sys.exit(1)

    keyring.set_password(SERVICIO, "USER_LPG", usuario_lpg)
    keyring.set_password(SERVICIO, "PASS_LPG", password_lpg)
    keyring.set_password(SERVICIO, "USER_LPE", usuario_lpe)
    keyring.set_password(SERVICIO, "PASS_LPE", password_lpe)

    keyring.set_password(SERVICIO, "TENANT_ID", tenant_id)
    keyring.set_password(SERVICIO, "CLIENT_ID", client_id)
    keyring.set_password(SERVICIO, "CLIENT_SECRET", client_secret)

    console.print(
            "\n[bold green]✔ [OK] Credenciales registradas exitosamente en el sistema.[/bold green]"
        )
    console.print(f"[dim]Namespace Regitrado :[/dim] [bold]{SERVICIO}[/bold]")

    time.sleep(1)
    sys.exit(0)
except KeyboardInterrupt:
    console.print(
        "\n\n[bold yellow]Operación interrumpida por el usuario.[/bold yellow]"
    )
    sys.exit(1)