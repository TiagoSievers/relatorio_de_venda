import pandas as pd
import win32com.client as win32

# faturamento por loja
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)
# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'tiagosievers@gmail.com'
mail.Subject = 'Relátorio de vendas'
mail.HTMLBody = f'''
Prezados

Segue o Relatório de vendas de cada loja.

Faturamento
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

Quantidade Vendida
{quantidade.to_html()}

Ticket Médio dos produtos de cada loja:
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

Qualquer dúvida estamos a disposição
att
Tiago Sievers
'''

mail.Send()