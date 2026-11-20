import os
from pydantic_settings import BaseSettings, SettingsConfigDict
from pydantic import field_validator
import keyring

AMBIENTE = os.getenv("APP_ENV", "DEV").upper()
SERVICIO = f"RPA_SUNAT_{AMBIENTE}"

class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file='.env', env_file_encoding='utf-8')

    SUNAT_RUC_LPG: str
    SUNAT_RUC_LPE: str
    PATH_CLOUD_TEST: str
    PATH_CLOUD: str
    
    SUNAT_USER_PROD_LPG: str
    SUNAT_PASSWORD_PROD_LPG: str
    SUNAT_USER_PROD_LPE: str
    SUNAT_PASSWORD_PROD_LPE: str

    TENANT_ID: str
    CLIENT_ID: str
    CLIENT_SECRET: str
    MAIL_BOX: str

    RECIPIENTS_MAIL_LIST: str | list[str]

    @field_validator("RECIPIENTS_MAIL_LIST", mode="after")
    @classmethod
    def parse_recipients(cls, v: str | list[str]) -> list[str]:
        if isinstance(v, str):
            correos = [
                c.strip() for c in v.replace("|", ",").split(",") if c.strip()
            ]
        else:
            correos = [str(c).strip() for c in v if str(c).strip()]

        if not correos:
            raise ValueError(
                "RECIPIENTS_MAIL_LIST debe contener al menos un correo válido."
            )
        return correos

    QUOTA_MORNING: int = 500
    QUOTA_AFTERNOON: int = 500
    QUOTA_NIGHT: int = 1000
    MAX_DAILY_QUOTA: int = 2000
    MAX_ZIP_SIZE_MB: float = 2.5

    SUNAT_USER_PROD_LPG = keyring.get_password(SERVICIO, "USER_LPG")
    SUNAT_PASSWORD_PROD_LPG = keyring.get_password(SERVICIO, "PASS_LPG")

    SUNAT_USER_PROD_LPE = keyring.get_password(SERVICIO, "USER_LPE")
    SUNAT_PASSWORD_PROD_LPE = keyring.get_password(SERVICIO, "PASS_LPE")

    TENANT_ID = keyring.get_password(SERVICIO, "TENANT_ID")
    CLIENT_ID = keyring.get_password(SERVICIO, "CLIENT_ID")
    CLIENT_SECRET = keyring.get_password(SERVICIO, "CLIENT_SECRET")
    
settings = Settings()