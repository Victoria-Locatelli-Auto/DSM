import json
from datetime import datetime

# Arquivo original que será sobrescrito
arquivo_json = "estado_autenticado.json"

# Nova data de expiração desejada (27 de maio de 2025 às 15:00)
nova_data_expiracao = datetime(2025, 5, 27, 15, 30)

# Converter para timestamp Unix
novo_timestamp = int(nova_data_expiracao.timestamp())

# Carregar JSON original
with open(arquivo_json, "r", encoding="utf-8") as f:
    dados = json.load(f)

# Alterar validade dos cookies que possuem `expires`
for cookie in dados.get("cookies", []):
    if cookie.get("expires", -1) > 0:  # Apenas cookies com validade definida
        print(f"🛠 Alterando validade do cookie '{cookie['name']}'")
        cookie["expires"] = novo_timestamp

# Salvar o arquivo atualizado, sobrescrevendo o original
with open(arquivo_json, "w", encoding="utf-8") as f:
    json.dump(dados, f, indent=4)

print(f"✅ Arquivo atualizado! Agora os cookies expiram em {nova_data_expiracao}.")