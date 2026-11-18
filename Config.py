import os
from pydantic_settings import BaseSettings, SettingsConfigDict
import keyring
# from pydantic import computed_field

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

    SUNAT_USER_PROD_LPG = keyring.get_password(SERVICIO, "USER_LPG")
    SUNAT_PASSWORD_PROD_LPG = keyring.get_password(SERVICIO, "PASS_LPG")

    SUNAT_USER_PROD_LPE = keyring.get_password(SERVICIO, "USER_LPE")
    SUNAT_PASSWORD_PROD_LPE = keyring.get_password(SERVICIO, "PASS_LPE")

    TENANT_ID = keyring.get_password(SERVICIO, "TENANT_ID")
    CLIENT_ID = keyring.get_password(SERVICIO, "CLIENT_ID")
    CLIENT_SECRET = keyring.get_password(SERVICIO, "CLIENT_SECRET")
    
settings = Settings()