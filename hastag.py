import pyautogui
import pyperclip
import time
import pandas as pd
import win32com.client as win
#import openpyxl
#import numpy


pyautogui.PAUSE=1

pyautogui.alert('PROGRAMA INICIADO, AGARDE')
pyautogui.press("win")
pyautogui.write("Edge")
pyautogui.press("enter")
time.sleep(2)
link="https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(2.75)
pyautogui.doubleClick(399,269)
time.sleep(1.5)
pyautogui.click()
pyautogui.click(1713,162)
time.sleep(0.8)
pyautogui.click(1555,537)
time.sleep(5)

tabela = pd.read_excel(r'E:\edge\Vendas - Dez.xlsx')
pd.set_option('display.max_columns', None)
faturamento =tabela['Valor Final'].sum()
quantidade =tabela['Quantidade'].sum()

outlook = win.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = f"curtishenrique10@gmail.com"
email.Subject = f"RELATORIO DE VENDAS"
email.HTMLBody = f"""
    <p><b>Olá, eu sou o Henrique: </b></p>
   <p> gostaria de te enviar o relatorio de vendas </p>

<p><b>Faturamento:
R${faturamento:,.2f}</b></p>

<p><b>Quantidade:
{quantidade:,}</b></p>

    """
print('EMAIL ENVIADO COM SUCESSO')
# enviar email
email.Send()
