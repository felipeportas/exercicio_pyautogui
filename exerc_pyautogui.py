#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Atualização de vendas do dia anterior.
# O seu trabalho diário é enviar um e-mail para a diretoria com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: felipeportas@hotmail.com
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# 

# In[4]:


import pyautogui as pyg
import time as tm
import pyperclip as pcp

pyg.PAUSE = 1
pyg.alert("Vai começar, aperte OK e não mexa em nada")


# ### Passo 1 - Abrir uma nova aba, entrar no link do sistema e abrir o arquivo

# In[5]:


pyg.hotkey('ctrl', 't')

link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pcp.copy(link)
pyg.hotkey("ctrl", "v")
pyg.press("enter")
tm.sleep(5)

# Acessar o arquivo
pyg.click (385, 290, clicks=2)
tm.sleep(3)
pyg.click (410, 286)
pyg.click (1091, 185)
tm.sleep(1)
pyg.click (899, 591)
tm.sleep(7)


# ### Passo 2 - Analisar o arquivo baixado (ler o arquivo e pegar os indicadores)

# In[6]:


import pandas as pd
tabela = pd.read_excel (r"E:\user\Download\Vendas - Dez.xlsx")

#lembrando que o Faturamento é igual à somatória da coluna Valor Final 
faturamento = tabela["Valor Final"].sum()

# e a Quantidade é a soma da coluna Quantidade
quantidade = tabela["Quantidade"].sum()

display (tabela)
print ("O faturamento é de", faturamento)
print ("A quantidade total é de", quantidade)


# ### Passo 3 - Enviar um e-mail pelo gmail

# In[8]:


# tem abrir uma nova aba
pyg.hotkey('ctrl', 't')
link1 = "mail.google.com"
pcp.copy(link1)
pyg.hotkey("ctrl", "v")
pyg.press("enter")
tm.sleep(3)
pyg.click (109, 195)
tm.sleep(3)

#inserindo e-mail
link2 = "felipeportas@hotmail.com"
pcp.copy(link2)
pyg.hotkey("ctrl", "v")

#inserindo o assunto
pyg.press("tab")
pyg.write ("Faturamento e Quantitativo Semanal")
pyg.press("tab")

#escrever o email

texto_email = f"""
Prezados, bom dia.

O faturamento semanal foi de R$ {faturamento:,.2f}, com a seguinte quantidade de produtos: {quantidade:,}.

Att,
Felipe
"""
pyg.write(texto_email)

#clicar em enviar - atalho do gmail
pyg.hotkey("ctrl", "enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[ ]:


time.sleep(5)
pyautogui.position()

