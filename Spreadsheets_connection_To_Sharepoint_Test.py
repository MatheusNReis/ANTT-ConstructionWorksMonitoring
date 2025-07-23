# Test the connection of spreadsheets stored in Sharepoint for proccesses automation
# Testa conexão de planilhas armazenadas em Sharepoint para automação de processos

import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from io import BytesIO

# Credenciais do SharePoint
site_url = ''
client_id = ''
client_secret = ''

# Caminhos dos arquivos no SharePoint
file_url_A = 'https://aaaa.sharepoint.com/:x:/r/sites/'
file_url_B = 'https://xxxx.sharepoint.com/:x:/r/sites/'

# Autenticação usando App-Only Flow
credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(site_url).with_credentials(credentials)

# Função para baixar o arquivo Excel do SharePoint
def download_excel(ctx, file_url):
    file = ctx.web.get_file_by_server_relative_url(file_url)
    file_object = BytesIO()
    file.download(file_object).execute_query()
    file_object.seek(0)  # Resetar o ponteiro do arquivo
    return file_object

# Baixar e carregar as planilhas
file_A = download_excel(ctx, file_url_A)
file_B = download_excel(ctx, file_url_B)

# Ler as planilhas com pandas
workbook_A = pd.read_excel(file_A, sheet_name='yyyy')
workbook_B = pd.read_excel(file_B, sheet_name='Planilha1')

# Ler o valor da célula C7 da planilha A
value = workbook_A.at[6, 'C']  # Índice 6 corresponde à linha 7 (0-indexed)

# Atribuir o valor à célula C2 da planilha B
workbook_B.at[1, 'C'] = value  # Índice 1 corresponde à linha 2 (0-indexed)

# Salvar as alterações na planilha B
path_to_save = r'C:\Users\matheus.reis\Desktop\workbook_B.xlsx'
with pd.ExcelWriter(path_to_save) as writer:
    workbook_B.to_excel(writer, sheet_name='Planilha1', index=False)

# Opcional: Carregar de volta para o SharePoint (se necessário)
# Usar a API do SharePoint ou bibliotecas como `shareplum` para fazer o upload do arquivo atualizado
