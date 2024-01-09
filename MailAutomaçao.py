
import pandas as pd
import win32com.client as win32
from datetime import date


# Base de dados
tabela_vendas = pd.read_excel ('vendas.xlsx')



# Visualizaçao dos dados
pd.set_option ('display.max_columns', None)


# Faturamento por loja
Faturamento = tabela_vendas [['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(Faturamento)

print('-'* 47)

# Qtd de prod vendidos por loja
Quantidade = tabela_vendas [['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(Quantidade)

print('-'* 47)

# ticket medio por produto em cada loja
Ticket_medio = (Faturamento['Valor Final']/ Quantidade['Quantidade']).to_frame()
Ticket_medio = Ticket_medio.rename(columns={0:'Ticket Medio'})
print(Ticket_medio)

# envio de email com o relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'brian.f.franca@gmail.com'
mail.Subject = 'Relatório de Vendas ' + str(date.today())

mail.HTMLBody = f''' 
<p>Prezados,</p>

<p>Segue o Relatorio de Vendas de cada Loja</p>

<p>Faturamento:</p>
{Faturamento. to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{Quantidade. to_html(formatters={'Quantidade': 'R${:,.2f}'.format})}

<p>Ticket Medio dos Produtos em cada Loja:</p>
{Ticket_medio. to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposiçao.</p>

<p>att.,

<p>Brian Farias França</p>
 '''


mail.Send()






