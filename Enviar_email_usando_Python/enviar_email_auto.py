import win32com.client as win32

#criando a integração

outlook = win32.Dispatch('outlook.application')

#criar email

email = outlook.CreateItem(0)

#configurar info do emial 

# email.To = 'Destino'
# email.Subject = 'assunto'
# email.HTMLBody = 'Corpo do email'

faturamento = 1500
qtd_Prod = 10
ticket_medio = faturamento / qtd_Prod

email.To = 'rochagabrielly80@gmail.com;alunoia16@gmail.com'
email.Subject = 'Email automatico em Python'
email.HTMLBody = f'''
<p>Ola, esse é o codigo Python</p>
<p>o a faturamento foi de R${faturamento}</p>
<p>vendemos {qtd_Prod} produtos</p>
<p>o ticket medio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Gabrielly</p>
'''

anexo = (r'C:\Users\gaby\Documents\email\media.xlsx')
email.Attachments.Add(anexo)
email.Send()
print("email enviado")