import pandas as pd
import plotly.express as px

tabela = pd.read_csv('telecom_users.csv')

pd.set_option('display.max_columns', None)
#print(tabela)

#excluir coluna // axis: COLUNA = 1, LINHA = 0
tabela = tabela.drop('Unnamed: 0',axis=1)
#print(tabela)

#printando as informações da tabela (objetc = texto, int = a numero inteiro e float = numeros com casas decimais
#print(tabela.info())

tabela['TotalGasto'] = pd.to_numeric(tabela['TotalGasto'], errors='coerce')

#tratamento de dados // dropna = exclue colunas vazias // 'all' = tudo e 'any' = pelo menos um valor vazio
tabela = tabela.dropna(how='all' , axis=1)

tabela = tabela.dropna(how='any', axis=0)

#analise inicial // value_counts = conta os valores de uma coluna
print(tabela['Churn'].value_counts(normalize = True).map('{:.1%}'.format))

#analise detalhada dos clientes // plotly.express = cria um grafico // o 'x' e o 'color' comparam uma coluna com a outra
grafico = px.histogram(tabela, x='Aposentado', color='Churn')
grafico.show()