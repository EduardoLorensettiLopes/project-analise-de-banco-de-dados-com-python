# ATENÇÃO Só funciona no meu pc para funcionar em outro alterar os pyautogui.click
import time
import pyautogui
import pyperclip
import pandas
import numpy
import openpyxl

pyautogui.PAUSE = 1
# passo 1: Entrar no sistema da empresa
pyautogui.press('win')
pyautogui.write('Google')
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey('win', 'up')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)

# passo 2: Navegar no sistema e encontrar a base de vendas
pyautogui.click(x=377, y=299, clicks=2)
time.sleep(2)
# passo 3: Fazer o download da base de vendas
pyautogui.click(x=401, y=373)
pyautogui.click(x=1694, y=178)
pyautogui.click(x=1439, y=625)
time.sleep(5)  # Esperar download acabar

# passo 4: Importar a base de vendas para o python
tabela = pandas.read_excel(r"C:\Users\Edu4r\Downloads\Vendas - Dez.xlsx")
print(tabela)
# passo 5: Calcular oo indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)
# passo 6: Enviar um e-mail para a diretoria com os indicadores de venda
# Abrir nova aba
pyautogui.press('win')
pyautogui.write('Google')
pyautogui.press('enter')
time.sleep(1)

# entrar no link email

pyautogui.hotkey('win', 'up')
pyperclip.copy('https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)

# clicar no botão escrever

pyautogui.click(x=74, y=188)
time.sleep(2)

# preencrer a informações do e-mail

pyperclip.copy('email de quem quer enviar')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
time.sleep(1)

pyautogui.write('Relatório de Vendas')

pyautogui.press('tab')
texto = f"""
Prezados,

Segue relatório de vendas
Faturamento: R${faturamento:,.2f}
Quantidades de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
att.,
Eduardo Lorensetti Lopes
"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

# enviar e-mail
pyautogui.hotkey('ctrl', 'enter')
