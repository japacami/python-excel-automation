# Biblioteca "pandas: permite o python e excel interagirem.
# "openpyxl" = open + python + excel
# Caso de filtrar colunas:
    # Utilizar "['']"

# Caso de agrupamento
    # utilizar .groupby()

# Casos de divisão, multiplicação, transformará em uma tabela:
    # .to_frame()

# "pywin32" = python + windows32

import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-' * 50)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)

# Quantidade de produtos vendidos por loja
qtd_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos)

print('-' * 50)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
print(ticket_medio)

print('-' * 50)

# Enviar um e-mail com o relatório
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = 'camila.sato@corujaconsultoria.com.br'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados, </p>

<p>Segue o relatório de vebdas por cada loja.</p>

<p>Faturamento: <p/>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida: <p/>
{qtd_produtos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja: <p/>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida fico à disposição.<p/>

<p>Atenciosamente, <p/>
'''

mail.Send()

print('E-mail enviado!')