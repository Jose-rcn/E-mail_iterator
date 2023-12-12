#import openpyxl
import smtplib

smtp_server = 'mail.iff.edu.br'
smtp_port = 587



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
    break