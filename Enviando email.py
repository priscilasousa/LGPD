import win32com.client as win32

# Criando a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# Criando email
email = outlook.CreateItem(0)

# Configurando as informações do email
email.To = "priscilla.sooousa@gmail.com"
email.Subject = "[ALERTA] Possível vazamento de Informações"
email.HTMLBody ="""
<p>Caro administrador de Segurança,</p>

<p>O usuário X está tentando enviar um email que pode conter informações sensíveis.</p>

<p>Por favor, verificar e tomar as ações necessárias!</p>
<p>Este é um email automático, por favor, não responder!</p>
"""
email.Attachments.Add()

email.Send()
print("Email enviado com sucesso!")
