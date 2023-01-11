import pyautogui as py
import pyperclip
import time
import pandas as pd


#passo 1 entrar no sistema
py.PAUSE = 1
py.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
py.hotkey("ctrl", "v")
py.press("enter")
time.sleep(3)

#passo 2 navegar até o local do arquivo
py.click(x=401, y=274, clicks=2)
time.sleep(1)

# passo 3 exportar relatório
py.click(x=390, y=332)
time.sleep(1)
py.click(x=1074, y=150)
time.sleep(1)
py.click(x=842, y=595)
time.sleep(3)

#passo 4 calcular indicadores

tabela = pd.read_excel(r'C:\Users\Lan\Downloads\Vendas - Dez.xlsx')
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)

#passo 5 mandar email
py.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
py.hotkey("ctrl", "v")
py.press("enter")
time.sleep(2)
py.click(x=89, y=172)
# preencher campos
# destinatario
py.write("klausmachado1@gmail.com")
py.press("enter")
py.press("tab")
# assunto
pyperclip.copy("relatório")
py.hotkey("ctrl", "v")
py.press("tab")
# texto
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de : R${faturamento:,.2f}
A quantidade de produtos foi de : R${quantidade:,.2f}

Att
Raphael
"""
pyperclip.copy(texto)
py.hotkey("ctrl", "v")
py.hotkey('ctrl', 'enter')