import pandas as pd
import win32com.client as win
#importando a base de dados
vendas = pd.read_excel('Vendas.xlsx')
#visualizar a base de dados
pd.set_option('display.max_columns', None)
#print(vendas)
#calcular o faturamento por loja
faturamento = vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
#print(faturamento)
#quantidade de produtos vendidos por lojas
produtos_vendidos = vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
#print(produtos_vendidos)
#print('-' *50)
#ticket medio por produto em cada loja // to_frame transfoma em tabela
ticket_medio = (faturamento['Valor Final'] / produtos_vendidos ['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Tícket Médio'})
#print(ticket_medio)
#enviar um email com o relatorio
outlook = win.Dispatch('outlook.application')
# Criar email
email = outlook.CreateItem(0)
# para enviar para mais de 1 pessoa colocar ";" e o email. // ex: curtishenrique10@gmail.com; teste1234@gmail.com
email.To = f"curtishenrique10@gmail.com"
email.Subject = f"RELATORIO DE VENDAS"
email.HTMLBody = f"""
    <p><b>Olá, eu sou o Henrique: </b></p>
   <p> gostaria de te enviar o relatorio de vendas </p>

<p><b>Faturamento:
{faturamento.to_html(formatters={'Valor Final': 'R${:,.0f}'.format})}</b></p>

<p><b>Quantidade:
{produtos_vendidos.to_html()}</b></p>

<p><b>Ticket medio dos produtos de cada loja:
{ticket_medio.to_html(formatters={'Tícket Médio': 'R${:,.2f}'.format})}</b></p>


    """
print('EMAIL ENVIADO COM SUCESSO')
# enviar email
email.Send()

