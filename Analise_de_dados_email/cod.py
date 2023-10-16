# importar a base de dados
# visializar a base de dados
# Faturamento por loja
# quantidade de produtos vendidos por loja
# ticket médio de produtos por loja
# enviar um email com o relatório


import pandas as pd
import win32com.client as win32

tab_vendas = pd.read_excel("Vendas.xlsx") 

#pd.set_option('display.max_columns', None) # Para  mostrar tudo da tabela - Dependendo da "IDE", não precisa 


#print(tab_vendas)

# Faturamento por loja

faturamento = tab_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() # mostrar o que eu quero, e tirar as duplicatas das lojas e sumar o  resto das tabela

print(faturamento)

# quantidade de produtos vendidos por loja

print('-' * 50)
quantidade = tab_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(quantidade)


# ticket médio por produto em cada loja
print('-' * 50)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()  # vai deixar em uma tabela, posso dar um nome a ela
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'}) # nomenado a coluna, ponho o nome dela atual para alterar

print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gamazzo0900@gmail.com'
mail.Subject = 'Teste testando sá Budega'
mail.HTMLBody = f'''
<p>Prezados,</p> 

<p>Segue o relatório de vendas por cada loja</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, pergunte pro estagiário que o cabra é bão.</p>

<p>Att.,</p>
<p>Gaio</p>

'''
# 3 aspas faz é pra textos com mais de uma linha no python 
mail.Send()

print('Email enviado')
