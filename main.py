import win32com.client as win32

# Criando a lista de destinatários e anexos
emails = [
    {"email":"joao.c.o.c11@gmail.com", "anexo":r"C:\Users\joaop\PycharmProjects\Email_Python\teste.txt"},
    {"email":"joao.p.s.s.8@gmail.com", "anexo":r"C:\Users\joaop\PycharmProjects\Email_Python\teste3.txt"}
]

# inicializando o outlook
outlook = win32.Dispatch('outlook.application')

# Loop para enviar os e-mails
for item in emails:
    mail = outlook.CreateItem(0)
    mail.To = item["email"]
    mail.Subject = "Provisão de Férias"
    mail.HTMLBody = f'''
            <p>Prezados,<p>
            <p>Olá, este é um e-mail automático enviado para {item["email"]}<p>
            <p>Abs,<p>
            <p>João Pedro<p>
'''
    # Adiciona o anexo
    for anexo in item["anexo"]:
        mail.Attachments.Add(item["anexo"])

    # Envia o e-mail
    mail.Send()

print("Todos os e-mails foram enviados com sucesso!")