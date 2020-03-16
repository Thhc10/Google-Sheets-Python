import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from tkinter import *

scope = ['https://www.googleapis.com/auth/spreadsheets']

# Arquivo JSON com as credenciais
credentials = ServiceAccountCredentials.from_json_keyfile_name('xxxxxx-1111.json', scope)

gc = gspread.authorize(credentials)

# Key da spreadsheet's url
wks = gc.open_by_key('xxxxxxxx') 

worksheet = wks.worksheet("Por Projeto") # Nome do worksheet

# Proxima Parcela

dia = worksheet.col_values(8)
mes = worksheet.col_values(9)
ano = worksheet.col_values(10)
parcelaspagas = worksheet.col_values(14)
parcelasrestantes = worksheet.col_values(16)

for i in range(1, len(dia)):

    if parcelasrestantes[i] == '0':  # Termino do Emprestimo
        continue

    dia_prox_parcela = dia[i]
    mes_prox_parcela = int(mes[i]) + 1 + int(parcelaspagas[i])
    ano_prox_parcela = 0

    if mes_prox_parcela > 24:
        mes_prox_parcela -= 24
        ano_prox_parcela = 2

    elif mes_prox_parcela > 12:
        mes_prox_parcela -= 12
        ano_prox_parcela = 1

    ano_prox_parcela += int(ano[i])

    a = str(dia_prox_parcela) + "/" + str(mes_prox_parcela) + "/" + str(ano_prox_parcela)
    worksheet.update_cell(i + 1, 19, a)

# Impressao do texto

texto_boleto = ""
texto_atraso = ""
msg = ""

for i in range(2, len(dia) + 1):

    data_alvo = worksheet.cell(i, 19).value
    data_hoje = datetime.today()
    data_hoje = data_hoje.strftime("%d/%m/%Y")

    data1 = datetime.strptime(data_alvo, "%d/%m/%Y")
    data2 = datetime.strptime(data_hoje, "%d/%m/%Y")

    delta_dias = (data1 - data2).days

    if delta_dias < 0:
        msg = "ATRASO da " + worksheet.cell(i, 1).value + " de " + str(abs(delta_dias)) + " dias.\n"
        dias_atraso = abs(delta_dias)
        worksheet.update_cell(i, 12, abs(delta_dias))
        texto_atraso += msg

    elif delta_dias <= 5:
        msg = "Boleto da " + worksheet.cell(i, 1).value + " em " + str(delta_dias) + " dias.\n"
        texto_boleto += msg

    msg = ""

texto = texto_boleto + "\n" + texto_atraso

# Interface
class Application:
    def __init__(self, master=None):
        self.widget1 = Frame(master)
        self.widget1.pack()
        self.msg = Label(self.widget1, text=texto, justify=LEFT)
        self.msg.pack()
root = Tk()
root.title('meBanq - Acompanhamento de Projetos')
root.iconbitmap('logo.ico')
Application(root)
root.mainloop()
