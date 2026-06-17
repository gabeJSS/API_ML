"""
Debug isolado para _buscar_frete_csv.
Rode com: python debug_frete.py
Edite as 3 variáveis abaixo antes de rodar.
"""

import json
import csv
import io
import requests
from pathlib import Path

# ─── EDITE AQUI ───────────────────────────────────────────────────────────────
COOKIE_PATH = Path("cookie.json")   # mesmo caminho que o core.py usa
ORDER_ID    = ""    # pedido que deveria ter frete
CLIENT_ID   = ""    # seu client.id (o que aparece no curl)
# ──────────────────────────────────────────────────────────────────────────────

def carregar_sessao():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    cookies = json.loads(COOKIE_PATH.read_text(encoding="utf-8"))
    for c in cookies:
        session.cookies.set(c["name"], c["value"], domain=c["domain"])
    return session

def obter_csrf(session):
    r = session.get("https://myaccount.mercadolivre.com.br/purchases/list", timeout=15)
    import re
    m = re.search(r'name="csrf-token"\s+content="([^"]+)"', r.text)
    return m.group(1) if m else None

session  = carregar_sessao()
csrf     = obter_csrf(session)
print(f"csrf_token: {csrf!r}\n")

params = {
    "requestInfo[method]": "GET",
    "requestInfo[path]": "/my_purchases/middleend/web/report/csv",
    "requestInfo[params][0][key]": "client.id",
    "requestInfo[params][0][value]": CLIENT_ID,
    "requestInfo[params][1][key]": "device.type",
    "requestInfo[params][1][value]": "desktop",
    "requestInfo[params][2][key]": "search.value",
    "requestInfo[params][2][value]": ORDER_ID,
}
headers = {
    "x-csrf-token": csrf,
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://myaccount.mercadolivre.com.br/my_purchases/list",
}

r = session.get(
    "https://myaccount.mercadolivre.com.br/my_purchases/api/web/download-csv",
    params=params, headers=headers, timeout=15,
)

print(f"Status: {r.status_code}")
print(f"Content-Type: {r.headers.get('Content-Type')}")
print(f"\n--- BODY BRUTO (primeiros 2000 chars) ---\n{r.text[:2000]}\n")

# Tenta parsear como CSV e mostra todas as linhas + colunas encontradas
print("--- PARSE CSV (todas as linhas) ---")
reader = csv.DictReader(io.StringIO(r.text), delimiter=";")
print(f"Colunas detectadas: {reader.fieldnames}\n")
for i, row in enumerate(reader):
    print(f"Linha {i+1}:")
    for k, v in row.items():
        print(f"  {k!r}: {v!r}")
    print()