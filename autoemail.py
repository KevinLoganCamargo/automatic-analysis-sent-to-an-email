import pyautogui
import pyperclip
from time import sleep
import pandas as pd
import time
import schedule

def DoTheWork():
  try:
    pyautogui.PAUSE=1
    pyautogui.press('win')
    pyautogui.write('Microsoft')#Utilizado para escrever algo
    pyautogui.press('enter')
    pyautogui.click(x=454, y=48)
    pyautogui.press('backspace')
    pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    sleep(4)
    pyautogui.click(x=747, y=408, clicks=2)
    sleep(2)
    pyautogui.click(x=655, y=501)
    pyautogui.click(x=1621, y=234)
    pyautogui.click(x=1333, y=787)
    sleep(2)
    tabela= pd.read_excel(r'C:\Users\kevin\Downloads\Vendas - Dez.xlsx')
    print(tabela)
    pyautogui.hotkey('ctrl','alt','tab')
    pyautogui.press('enter')
    faturamento=tabela['Valor Final'].sum()
    quantidade_produtos=tabela['Quantidade'].sum()
    sleep(4)
    pyautogui.hotkey('ctrl','esc')
    pyautogui.write('Microsoft')
    pyautogui.press('enter')
    pyautogui.write('https://mail.google.com/mail/u/0/?hl=pt-BR#inbox')
    pyautogui.press('enter')
    sleep(4)
    pyautogui.click(x=165, y=269)
    sleep(4)
    pyautogui.write('kevinlogancamargo77@gmail.com')
    pyautogui.press('enter')
    pyautogui.press('tab')
    pyautogui.write('Analise Simples')
    pyautogui.press('tab')
    pyautogui.write('Prezados, bom dia\n\nO faturamento de ontem foi de:R${:,.2f}\nA quantidade de produtos vendidos foi de:{:,.2f}\n\nAbs\nKevinLogan'.format(faturamento,quantidade_produtos))
    pyautogui.press('tab')
    pyautogui.press('enter')
  except: print('Infelizmente ocorreu algum erro inesperado')

schedule.every().tuesday.at('17:31').do(DoTheWork)#Used to do a task run at certain time, day and constant

while 1:
 schedule.run_pending()








