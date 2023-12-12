#import openpyxl
import smtplib

smtp_server = 'mail.iff.edu.br'
smtp_port = 587

def connect_server(server, port):
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
    return True

while True:
    try:
        print(connect_server(smtp_server,smtp_port))
    except Exception:
        print("Erro ao se conectar com o servidor. Verifique as informações do servidor")
        break
    print('{:-^40}'.format('Log in'))
    user_username = input("Informe o seu endereço de e-mail: ")
    user_password = input('Informe a senha: ')
    print('-'*40)

    server.login(username, password)