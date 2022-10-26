import pyautogui
import time
import pyperclip
#permite controlar o mouse,teclado e tela
#escreverpyautogui  pip install pyautogui
import pyperclip
import time
pyautogui.PAUSE= 3

#pyautogui.click
#pyautogui.write
#pyautogui.press
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("Enter")

pyautogui.hotkey("ctrl","t")
time.sleep(3)
pyperclip.copy("https://drive.google.com/drive/u/0/mobile/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl","v")
pyautogui.press("Enter")
time.sleep(5)
pyautogui.click(x=569, y=501,clicks=2)
time.sleep(2)
pyautogui.click(x=698, y=412,clicks=2)
time.sleep(9)
tabela= pandas.read_excel(r"C:\Users\marco\Downloads\Vendas - Dez.xlsx")
display(tabela)
quantidade=tabela['Quantidade'].sum()
faturamento = tabela['Valor Final'].sum()
  print(faturamento)
    print(quantidade)
pyautogui.hotkey("ctrl","t")
pyautogui.write("mail.google.com/mail/u/#inbox")
pyautogui.press("enter")
sleep(8)
#mandar email
#clicarbotao +
pyautogui.click(x=102, y=166,clicks=2)
pyautogui.write("schuck.2241@gmail.com")
pyautogui.press("tab")
sleep(1)

#escrever assunto
pyperclip.copy("relatorio de vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
#escrever email
texto=f""" prezados,
o faturamento foi de {faturamento} e a quantidade foi de {quantidade}"""
pyautogui.write(texto)