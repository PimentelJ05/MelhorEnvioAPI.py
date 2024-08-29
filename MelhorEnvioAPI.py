import requests
from openpyxl import Workbook

# Configurações da API
AMBIENTE = 'sandbox'  # Alterar para 'producao' se estiver usando o ambiente de produção

if AMBIENTE == 'producao':
    API_URL = 'https://melhorenvio.com.br/v2/orders'  # URL de produção
else:
    API_URL = 'https://sandbox.melhorenvio.com.br/v2/orders'  # URL de sandbox

API_KEY = 'seu_token_de_acesso_aqui'  # Substitua pelo seu token de acesso


# Função para obter dados dos pedidos
def obter_dados_pedidos():
    headers = {
        'Authorization': f'Bearer {API_KEY}',
        'Content-Type': 'application/json'
    }

    try:
        response = requests.get(API_URL, headers=headers)
        response.raise_for_status()  # Verifica se houve um erro na solicitação
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        print(f"Erro HTTP ao fazer a solicitação: {http_err}")
    except Exception as err:
        print(f"Erro inesperado ao fazer a solicitação: {err}")


# Função para salvar os dados em uma nova planilha
def salvar_em_planilha(dados):
    if not dados:
        print("Nenhum dado foi obtido da API.")
        return

    # Cria um novo workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Pedidos'

    # Cabeçalhos
    sheet['A1'] = 'Nome do Cliente'
    sheet['B1'] = 'ID do Pedido'
    sheet['C1'] = 'Status do Pedido'

    # Adiciona os dados
    for i, pedido in enumerate(dados, start=2):  # Supondo que 'dados' seja uma lista de pedidos
        sheet[f'A{i}'] = pedido.get('customer', {}).get('name', 'N/A')  # Exemplo de como acessar o nome do cliente
        sheet[f'B{i}'] = pedido.get('id', 'N/A')
        sheet[f'C{i}'] = pedido.get('status', 'N/A')

    # Salva a planilha na pasta desejada
    caminho_arquivo = 'G:/Drives compartilhados/pedidos.xlsx'
    workbook.save(caminho_arquivo)
    print("Planilha criada e salva com sucesso!")


def main():
    try:
        dados_pedidos = obter_dados_pedidos()
        salvar_em_planilha(dados_pedidos)
        print("Planilha atualizada com sucesso!")
    except Exception as e:
        print(f"Erro ao atualizar a planilha: {e}")


if __name__ == "__main__":
    main()
