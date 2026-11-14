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

    SUNAT_USER_PROD_LPG = keyring.get_password(SERVICIO, "USER_LPG")
    SUNAT_PASSWORD_PROD_LPG = keyring.get_password(SERVICIO, "PASS_LPG")

    SUNAT_USER_PROD_LPE = keyring.get_password(SERVICIO, "USER_LPE")
    SUNAT_PASSWORD_PROD_LPE = keyring.get_password(SERVICIO, "PASS_LPE")
    
settings = Settings()