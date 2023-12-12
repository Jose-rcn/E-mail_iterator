import openpyxl
import smtplib
from pandas import ExcelFile


smtp_server = 'mail.iff.edu.br'
smtp_port = 587

def select_sheet(table):
    table_sheets = table.sheet_names
    for index,sheet in enumerate(table_sheets):
        print(f'[{index}] {sheet}')
    try:
        selected_sheet  = int(input("Selecione a pagina da tabela: "))
        for index,sheet_name in enumerate(table_sheets):
            if index == selected_sheet:
                return sheet_name
    except TypeError:
        print('Digite um número!')
    except IndexError:
        print('Digite um número válido')
def sheet_connect(sheet_name):
    table = openpyxl.load_workbook(table_path)
    return table[sheet_name]
def column_to_number(column):
    column = column.upper()
    if 'A' <= column <= 'Z':
        ordem_numerica = ord(column) - ord('A') + 1
        return ordem_numerica


while True:
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        print('Conectado ao servidor.')
    except Exception:
        print("Erro ao se conectar com o servidor. Verifique as informações do servidor")
        break
    print('{:-^40}'.format('Log in'))
    user_username = input("Informe o seu endereço de e-mail: ")
    user_password = input('Informe a senha: ')
    print('-'*40)
    try:
        server.login(user_username, user_password)
        print('Usuário logado')
    except smtplib.SMTPAuthenticationError:
        print('E-mail ou senha incorretos')
        continue
    table_path = input('Informe o caminho para acessar o arquivo: ')
    table_path = 'C:/Users/joser/OneDrive/Documentos/MeusRepositórios/E-mail_iterator/Test_table.xlsx'
    sheet_name = select_sheet(ExcelFile(table_path))
    while sheet_name is None:
        sheet_name = select_sheet(ExcelFile(table_path))
    sheet = sheet_connect(sheet_name)

    email_column = input('Informe a coluna da planilha que possui os endereços de e-mail: ')
    sheet.iter_rows(min_row=2,values_only=True)
    
    break