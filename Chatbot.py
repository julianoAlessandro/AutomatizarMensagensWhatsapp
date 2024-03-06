# Descrever os passos manuais e dps transformar isso em codigo
# 1- Abrir a planilha do excel.
# 2- pegar o nome e o telefone com vencimento.
# 3 - Criar links personalizados do whatsaap e enviar mensagens para cada cliente com base nos dados da planilha
#  link personalizado whatsapp
# clicar em enviar
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui as py

webbrowser.open("https://web.whatsapp.com/")
sleep(30)
workbook = openpyxl.load_workbook("Pessoa.xlsx")

ContatoPessoas = workbook["ListaPessoas"]

for linha in ContatoPessoas.iter_rows(min_row=2, max_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    Data = linha[2].value
    mensagem = f"Ola senhor {nome}, sua fatura venceu recentemento do cartão nubaank na data:{Data.strftime("%d/%m/%Y")}"

    link_mensagem = (
        f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
    )
    webbrowser.open(link_mensagem)
    sleep(12)
    try:
        sleep(5)
        py.moveTo(1326, 690, duration=2)
        py.click()
        sleep(5)
        py.moveTo(744, 194, duration=2)
        py.click()
        py.hotkey("ctrl", "w")
        sleep(5)

    except:
        print(f"nao foi possivel enviar mensagem")
        with open("erros.csv", "a", newline="", encoding="utf-8") as arquivo:
            arquivo.write(f"{nome},{telefone}")
py.hotkey("ctrl", "w")
