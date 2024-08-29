import requests
from openpyxl import Workbook
import json

# Novo token de acesso obtido
token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxNTg5NyIsImp0aSI6IjRiNGU2ZmUwMTljMTU0ZmU1MGEwODI1ZDZkMDgzMTVhMzJkNTA1ZDBhY2IwYjkwMWU5YWU5YmFiNzU2MjRiNzhhYjRjMDlmNjhkYjQ0OTFkIiwiaWF0IjoxNzI0OTU1MjI4Ljk4MTIwNSwibmJmIjoxNzI0OTU1MjI4Ljk4MTIwNywiZXhwIjoxNzI3NTQ3MjI4Ljk0ODQ0Niwic3ViIjoiIiwic2NvcGVzIjpbXX0.W_w_VN5UjwBCwEc3amWN9g_sFEa4kwvOYDkV9GQhJA38AOXU6yHh7uubI486_j3Pwg-G5B-g6H0M68VJWSZFgU7kZmeXL1PbPkE8Xmg44HHoXtWhi4k1YnzcITiMDMA0Zb9c44RH9QgdzfKYg8EB69650-49hKJmv5dxCUftNUveBjc9QxerdZ3SnMPE4k1D4goz481Tdd1RuaxFVBZYto7fmZaEyeGfFSw44shvfIInA51S369e41LZeXvemc0RjbnOaKu5XbNjjzH9IcWElby0rzrwYFb4B6Jb4xx94JWQi0ZwWL-JoDw9ZwuJ6IrdwhtCaWzBcNDtihr8YLmWuKy3QI-V85zOEwOpL9c-2k0JFOlWMl6GE46QYSOSHBCatVx5Ixvmg1EzslUVm86acpouYQRcintTAm61HtnvsfEpiJnBtbcM8xJ7jWGGGbY0oXoXzELPenYOTzjrlTU1PdrT4XZedbc5ijkSJ9jXVILWRvp-kAexPdkCsbdeZEhSO6zSCvYTd4z85YrXs_cKFIE33BMuDmArR6fWU-l9jUb__q8KifR63HqThJo1I2Kdv6fbcCD5xzZtEcaj3ilTdPg-5kx-JAyXJTXlUKtppPKYv6iT_tne0VXczeQ0JmcCcYlVsRyz8zTfIHrjAwS9436jOoG1w3i6sar4G9LJ-Fw'

# Função para obter os pedidos
def obter_pedidos():
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'User-Agent': 'Planilha Crédito Essencial/1.0 (julia.pimentel@creditoessencial.com.br)'
    }
    url = 'https://api.melhorenvio.com.br/v2/me/shipment-orders'  # URL correta do endpoint de pedidos

    response = requests.get(url, headers=headers)

    # Verifica o status da resposta
    print(f"Status Code: {response.status_code}")
    print(f"Response Text: {response.text}")

    try:
        response.raise_for_status()  # Verifica se houve um erro na solicitação

        # Verifica se o conteúdo da resposta é JSON válido
        try:
            dados_json = response.json()
            return dados_json
        except json.JSONDecodeError:
            print("Erro ao analisar JSON: A resposta não é um JSON válido.")
            return None

    except requests.exceptions.HTTPError as http_err:
        print(f"Erro HTTP ao fazer a solicitação: {http_err}")
    except Exception as err:
        print(f"Erro inesperado ao fazer a solicitação: {err}")

# Função para salvar os dados dos pedidos em uma planilha
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
    for i, pedido in enumerate(dados.get('orders', []), start=2):  # Corrigido para acessar a lista correta
        sheet[f'A{i}'] = pedido.get('buyer', {}).get('name', 'N/A')
        sheet[f'B{i}'] = pedido.get('id', 'N/A')
        sheet[f'C{i}'] = pedido.get('status', 'N/A')

    # Salva a planilha na pasta desejada
    caminho_arquivo = 'GC:\Users\usuario/pedidos.xlsx'
    workbook.save(caminho_arquivo)
    print("Planilha criada e salva com sucesso!")

def main():
    try:
        dados_pedidos = obter_pedidos()
        salvar_em_planilha(dados_pedidos)
        print("Planilha atualizada com sucesso!")
    except Exception as e:
        print(f"Erro ao atualizar a planilha: {e}")

if __name__ == "__main__":
    main()
