#!pip install pyautogui
#!pip install pyperclip
#!pip install pandas
#As coordenadas do mouse são diferentes de computador p/ computador;
#Código feito no intensivão do canal Hashtag Programação

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 0.5
# Passo 1: Entrar no sistema da empresa (no nosso caso, o link do drive)
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5) #Aguardar o site carregar, utilizando a biblioteca TIME

# Passo 2: Navegar no sistema e encontrar o banco de dados (entrar na pasta exportar)
pyautogui.click(x=375, y=260, clicks=2)
time.sleep(4)

# Passo 3: Exportar/Fazer Download da base de dados
pyautogui.click(x=384, y=335)
pyautogui.click(x=1157, y=155)
pyautogui.click(x=947, y=565)
time.sleep(4)

import pandas as pd

# Passo 4: Importar a base de dados para o python
base_de_dados = pd.read_excel(r"C:\Users\jorge\Downloads\Vendas - Dez.xlsx")# na frente do endreço do CAMINHO, SEMPRE coloque um r
print(base_de_dados)

# Vamos agora ler o arquivo baixado para pegar os indicadores
# Faturamento ; Quantidade de Produtos
# Passo 5: Calcular os indicadores

# faturamento - soma da coluna valor final
faturamento = base_de_dados["Valor Final"].sum() #.sum - soma a coluna inteira
# quantidade - soma da coluna quantidade
quantidade = base_de_dados["Quantidade"].sum() #.sum - soma a coluna inteira

# Vamos agora enviar um e-mail pelo gmail
# Passo 6: Enviar um email para a diretoria com o relatório

# Abrir o email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
# Clicar no botão escrever
pyautogui.click(x=85, y=165)
time.sleep(6)
# Escrever o destinatario
pyautogui.write("jorgegbsouza@gmail.com")
pyautogui.press("tab") # Selecionar o email digitado
pyautogui.press("tab") # Passar para o próximo campo
# Escrevendo o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # Passar para o próximo campo
# Escrevendo o corpo do email
texto = f"""Prezados, Bom dia

O faturamento de ontem foi de R${faturamento:,.2f} reais
A quantidade foi de {quantidade:,}

Abs
Jorge Gabriel
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
# Enviar o email
pyautogui.hotkey("ctrl", "enter")