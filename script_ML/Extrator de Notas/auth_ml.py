import json
import requests
from datetime import datetime, timedelta
from config import CONFIG_PATH, AUTH_PATH

ML_TOKEN_URL   = "https://api.mercadolibre.com/oauth/token"
ML_USER_ME_URL = "https://api.mercadolibre.com/users/me"
_EXPIRY_MARGIN = timedelta(minutes=5)

DEFAULT_CONFIG = {"redirect_uri": "", "client_id": "", "client_secret": ""}
DEFAULT_AUTH   = {"access_token": "", "refresh_token": "", "seller_id": "", "expires_at": ""}

def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return {**DEFAULT_CONFIG, **json.loads(CONFIG_PATH.read_text(encoding="utf-8"))}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)

def save_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

def load_auth() -> dict:
    if AUTH_PATH.exists():
        try:
            return {**DEFAULT_AUTH, **json.loads(AUTH_PATH.read_text(encoding="utf-8"))}
        except Exception:
            pass
    return dict(DEFAULT_AUTH)

def save_auth(auth: dict):
    AUTH_PATH.write_text(json.dumps(auth, indent=2, ensure_ascii=False), encoding="utf-8")

def clear_auth():
    if AUTH_PATH.exists():
        AUTH_PATH.unlink()

def token_is_valid(auth: dict) -> bool:
    if not auth.get("access_token") or not auth.get("expires_at"):
        return False
    try:
        expiry = datetime.fromisoformat(auth["expires_at"])
        return datetime.now() < expiry - _EXPIRY_MARGIN
    except ValueError:
        return False

def exchange_code_for_token(cfg: dict, code: str) -> dict:
    resp = requests.post(ML_TOKEN_URL, data={
        "grant_type": "authorization_code", "client_id": cfg["client_id"],
        "client_secret": cfg["client_secret"], "code": code, "redirect_uri": cfg["redirect_uri"]
    }, timeout=15)
    if resp.status_code != 200:
        raise RuntimeError(f"Erro OAuth ({resp.status_code}): {resp.text[:200]}")
    
    data = resp.json()
    expires_in = int(data.get("expires_in", 21600))
    return {
        "access_token": data["access_token"],
        "refresh_token": data["refresh_token"],
        "seller_id": str(data.get("user_id", "")),
        "expires_at": (datetime.now() + timedelta(seconds=expires_in)).isoformat(timespec="seconds")
    }

def refresh_access_token(cfg: dict, auth: dict) -> dict:
    resp = requests.post(ML_TOKEN_URL, data={
        "grant_type": "refresh_token", "client_id": cfg["client_id"],
        "client_secret": cfg["client_secret"], "refresh_token": auth["refresh_token"]
    }, timeout=15)
    if resp.status_code != 200:
        raise RuntimeError(f"Erro ao renovar token: {resp.text[:200]}")
    
    data = resp.json()
    expires_in = int(data.get("expires_in", 21600))
    return {
        **auth,
        "access_token": data["access_token"],
        "refresh_token": data.get("refresh_token", auth["refresh_token"]),
        "expires_at": (datetime.now() + timedelta(seconds=expires_in)).isoformat(timespec="seconds")
    }

def ensure_valid_token(cfg: dict, auth: dict) -> dict:
    if token_is_valid(auth):
        return auth
    if not auth.get("refresh_token"):
        raise RuntimeError("Token expirado e sem refresh_token. Re-autentique.")
    new_auth = refresh_access_token(cfg, auth)
    save_auth(new_auth)
    return new_auth

def fetch_seller_id(access_token: str) -> str:
    resp = requests.get(ML_USER_ME_URL, headers={"Authorization": f"Bearer {access_token}"}, timeout=10)
    if resp.status_code != 200:
        raise RuntimeError(f"Erro ao buscar seller_id")
    return str(resp.json()["id"])