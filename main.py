import requests
import win32com.client as win32
import json

cotacoes = requests.get("https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL")
cotacoes = cotacoes.json()
cotacoes_dolar = cotacoes['USDBRL']['bid']
cotacoes_euro = cotacoes['EURBRL']['bid']
print("A cotação do dolar é: " + cotacoes_dolar)
print("A cotação do euro é: " + cotacoes_euro)

#envio de email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Your email here...'
mail.Subject = 'Relatório com as cotações'
mail.HTMLBody = f'''
<p>Prezado Anderson,</p> 

<p> <strong> Segue o relatório com as cotações atuais.</strong> </p> 


<p> <strong> Cotação Euro: </strong> </p>
{cotacoes_euro}


<p> <strong> Cotação Dolar: </strong> </p>
{cotacoes_dolar}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.,</p> 
<p>Kelvin</p>
'''

mail.Send()

print('\n Email enviado!')
