import requests
import json

def buscar_compra_ml(token_acesso, id_compra):
    url = f"https://api.mercadolibre.com/orders/{id_compra}"
    
    headers = {
        "Authorization": f"Bearer {token_acesso}"
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status() # Dispara erro se o token for inválido ou o ID não existir
        
        dados = response.json()
        return dados
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Erro na requisição: {e}")
        # Se a API retornou alguma mensagem de erro detalhada, ele imprime aqui
        if 'response' in locals() and response.text:
            print(f"Detalhes do erro: {response.text}")
        return None

# ==========================================
# TESTANDO O SCRIPT
# ==========================================
if __name__ == "__main__":
    TOKEN = ""
    ID_PEDIDO = ""
    
    resultado = buscar_compra_ml(TOKEN, ID_PEDIDO)
    
    if resultado:
        print("Informações da compra capturadas com sucesso:\n")
        # Imprime o JSON formatado e indentado para facilitar a leitura
        print(json.dumps(resultado, indent=4, ensure_ascii=False))