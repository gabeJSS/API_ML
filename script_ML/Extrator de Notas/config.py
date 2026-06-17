import os
from pathlib import Path

APP_NAME = "MLERPExtractor"

def get_appdata_dir() -> Path:
    base = os.getenv("APPDATA") or str(Path.home())
    d = Path(base) / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d

# ── Arquivos de Configuração e Autenticação (AppData)
CONFIG_PATH = get_appdata_dir() / "config.json"
AUTH_PATH   = get_appdata_dir() / "auth.json"
COOKIE_PATH = get_appdata_dir() / "cookie.json"

# ── Arquivos de Dados (Pasta do Script)
BASE_DIR     = Path(__file__).parent
JSON_EXTRATO = BASE_DIR / "../integracao_erp.json"
JSON_ENRIQ   = BASE_DIR / "../integracao_erp_enriquecido.json"
JSON_FINAL   = BASE_DIR / "../integracao_erp_final.json"
EXCEL_CLASS  = BASE_DIR / "../classificacao.xlsx"
XML_DIR      = BASE_DIR / "../xmls_nfe"