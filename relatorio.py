from datetime import datetime
from tkinter import Label
from tkinter.constants import W
from PIL.Image import new
import pyautogui
import time
import pyperclip
import pandas as pd

import os
from IPython.display import display
from traitlets.traitlets import Undefined

from gerais import *
from configuracoes import *
from usuario import *

import matplotlib.pyplot as plt
from matplotlib import rcParams
from fpdf import FPDF
import numpy as np
import matplotlib.cbook as cbook
from fpdf import FPDF, HTMLMixin
from datetime import datetime, timedelta
import zipfile

BDconfiguracoes = recuperaConfiguracoes()
BDusuariosRelatorio = {}
recuperaUsuarios(BDusuariosRelatorio)

# =====================================Variaveis globais à serem usadas============================
path = os.getcwd()
datasMinCasos = []
rcParams['figure.figsize'] = 10, 6
datasMesRetrasado = []
dadosCausasMesRetrasado = []
dadosObitosMesRetrasado = []
datasMesPassado = []
dadosCausasMesPassado = []
dadosObitosMesPassado = []
datasMaxCasos = []
tabelaDatasMin = """"""
tabelaDatasMax = """"""
# ======================================================================================================

# ====Verifica se relatorio existe====
def existeRelatorio(dic,chave):
    if chave in dic.keys():
        return True
    else:
        return False

# ====Insere relatorio==== +++++++++++++++++++++++++++++++++++++++++++++DADOS PROVISORIOS PARA BAIXO
def insereRelatorio(dic, dado1, dado2):
    data = dataAtualF()
    confirma = input(f"Deseja usar a data {data}? (S/N): ").upper()
    if confirma == 'N':
        data = input('Informe a data (dd/mm/yyyy): ')
    if existeRelatorio(dic, data):
        print("Relatorio já cadastrado!")
        pausa()
    else:
        dic[data]=(dado1, dado2)
        print("Dados inseridos com sucesso!")
        pausa()

# ====Exibe um relatorio====
def mostraRelatorio(dic,chave):
    if existeRelatorio(dic,chave):
        dados = dic[chave]
        print(f"Data: {chave}")
        print(f"Dado1: {dados[0]}")
        print(f"Dado2: {dados[1]}")
    else:
        print("Relatorio não cadastrada!")

# ====Altera relatorio====
def alteraRelatorio(dic,chave, dado1, dado2):
    if existeRelatorio(dic,chave):
        mostraRelatorio(dic,chave)
        confirma = input("Tem certeza que deseja alterá-lo? (S/N): ").upper()
        if confirma == 'S':
            dic[chave]=(dado1, dado2)
            print("Dados alterados com sucesso!")
            pausa()
        else:
            print("Alteração cancelada!")
            pausa()
    else:
        print("Usuario não cadastrado!")
        pausa()

# ====Remove um relatorio====
def removeRelatorio(dic,chave):
    if existeRelatorio(dic,chave):
        mostraRelatorio(dic,chave)
        confirma = input("Tem certeza que deseja apagar? (S/N): ").upper()
        if confirma == 'S':
            del dic[chave]
            print("Dados apagados com sucesso!")
            pausa()
        else:
            print("Exclusão cancelada!")
            pausa()
    else:
        print("Relatorio não cadastrado!")
        pausa()

# ====Mostra todos relatorios====
def mostraTodosRelatorios(dic):
    print("Relatório: Todas os relatorios\n")
    print("DATA - DADO1 - DADO2\n")
    for data in dic:
        tupla = dic[data]
        linha = data + " - " + tupla[0] + " - " + tupla[1]
        print(linha)
    print("")
    pausa()

# ======Grava dados no arquivo=====
def gravaRelatorios(dic):
    arq = open("relatorio.txt", "w")
    for data in dic:
        tupla = dic[data]
        linha = data+";"+tupla[0]+";"+tupla[1]+"\n"
        arq.write(linha)
    arq.close()

# ======Pega dados do arquivo====
def recuperaRelatorios(dic):
    if (existe_arquivo("relatorio.txt")):
        arq = open("relatorio.txt", "r")
        for linha in arq:
            linha = linha[:len(linha)-1]
            lista = linha.split(";")
            data = lista[0]
            dado1 = lista[1]
            dado2 = lista[2]
            dic[data] = (dado1, dado2)

def baixarArquivos():
    BDconfiguracoes = recuperaConfiguracoes()
    recuperaUsuarios(BDusuariosRelatorio)
    pyautogui.PAUSE = 1

    pyautogui.alert('O programa vai comecar, nao use nada do seu computador!')

    pyautogui.press('win')
    time.sleep(0.5)
    link = "https://drive.google.com/drive/u/0/folders/1dSPGvocU1Tp-KobogpWfCBJF8QXEcSCY"
    pyperclip.copy(link)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.click(int(BDconfiguracoes[0]), int(BDconfiguracoes[1]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[2]), int(BDconfiguracoes[3]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[4]), int(BDconfiguracoes[5]), clicks=1)
    time.sleep(5)
    pyautogui.click(int(BDconfiguracoes[8]), int(BDconfiguracoes[9]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[2]), int(BDconfiguracoes[3]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[4]), int(BDconfiguracoes[5]), clicks=1)
    time.sleep(5)
    # pyautogui.hotkey('ctrl', 'w')
    time.sleep(2)

def grafico1(tabela):
    plt.plot(tabela["Data"], tabela["Total de casos"], label="Total")
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"/GraficosRelatorios/totalDeCasosEstado-"+dataAtualFArquivo()+".png")
    plt.close()

def grafico2(tabela):
    plt.plot(tabela["Data"], tabela["Casos por dia"], label="Casos")
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"/GraficosRelatorios/casosPorDiaEstado-"+dataAtualFArquivo()+".png")
    plt.close()

def grafico3(tabela):
    plt.plot(tabela["Data"], tabela["Óbitos por dia"], label="Óbitos")
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"/GraficosRelatorios/obitosPorDiaEstado-"+dataAtualFArquivo()+".png")
    plt.close()

def grafico4(tabela):
    plt.plot(tabela["Data"], tabela["Total de casos"], label="Total")
    plt.plot(tabela["Data"], tabela["Casos por dia"], label="Casos")
    plt.plot(tabela["Data"], tabela["Óbitos por dia"], label="Óbitos")
    plt.ylim(0, 30000)
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"/GraficosRelatorios/dadosEstado-"+dataAtualFArquivo()+".png")
    plt.close()

def grafico5(tabela, listaDatas):
    mesPassado = int(mesAtual())
    mesRetrasado = int(mesAtual())
    if mesPassado == 2:
        mesPassado=1
        mesRetrasado=12
    elif mesPassado == 1:
        mesPassado = 12
        mesRetrasado = 11
    else: 
        mesPassado-=1
        mesRetrasado-=2
    continua = True
    x = int(len(listaDatas))-1
    while continua==True:
        dataFra = (str(listaDatas[x]).split(' ')[0].split('-'))
        if int(dataFra[1]) == mesPassado:
            datasMesPassado.append(tabela["Data"][x])
            dadosCausasMesPassado.append(int(tabela["Casos por dia"][x]))
            dadosObitosMesPassado.append(int(tabela["Óbitos por dia"][x]))
        if int(dataFra[1]) == mesRetrasado:
            datasMesRetrasado.append(tabela["Data"][x])
            dadosCausasMesRetrasado.append(int(tabela["Casos por dia"][x]))
            dadosObitosMesRetrasado.append(int(tabela["Óbitos por dia"][x]))
        if x==0 or int(dataFra[1]) < mesRetrasado:
            continua = False
        else:
            x-=1
        
    labels = ['Casos', 'Óbitos']
    retrasado = [sum(dadosCausasMesRetrasado), sum(dadosObitosMesRetrasado)]
    passado = [sum(dadosCausasMesPassado), sum(dadosObitosMesPassado)]

    x = np.arange(len(labels))  # the label locations
    width = 0.35  # the width of the bars

    fig, ax = plt.subplots()
    rects1 = ax.bar(x - width/2, retrasado, width, label='Mês retrasado')
    rects2 = ax.bar(x + width/2, passado, width, label='Mês passado')

    # Add some text for labels, title and custom x-axis tick labels, etc.
    ax.set_ylabel('Quantidade')
    ax.set_title('Casos e óbitos - mes retrasado/passado')
    ax.set_xticks(x, labels)
    ax.legend()

    ax.bar_label(rects1, padding=3)
    ax.bar_label(rects2, padding=3)

    fig.tight_layout()

    plt.savefig(path+"/GraficosRelatorios/dadosComparativos-"+dataAtualFArquivo()+".png")
    plt.close()

def gerarTabelas(indicesMin, listaDatas, minCasos, indicesMax, maxCasos):
    tabelaDatasMin = """
        <font size="15"><p align=center>Datas de menor ocorrencia de casos</p></font>
        <table width="50%">
            <thead>
                <tr>
                <th width="20%">#</th>
                <th width="50%">Data</th>
                <th width="20%">Total de casos</th>
                </tr>
            </thead>
            <tbody>
    """

    tabelaDatasMax = """
        <font size="15"><p align=center>Datas de maior ocorrencia de casos</p></font>
        <table width="50%">
            <thead>
                <tr>
                <th width="20%">#</th>
                <th width="50%">Data</th>
                <th width="20%">Total de casos</th>
                </tr>
            </thead>
            <tbody>
    """
    

    for indMin in range(len(indicesMin[0])):
        datasMinCasos.append(listaDatas[indicesMin[0][indMin]])
        tabelaDatasMin+="""
            <tr>
                <td align=center>{}</td>
                <td align=center>{}</td>
                <td align=center>{:.1f}</td>
            </tr>
        """.format(indMin+1, listaDatas[indicesMin[0][indMin]], minCasos)

    for indMax in range(len(indicesMax[0])):
        datasMaxCasos.append(listaDatas[indicesMax[0][indMax]])
        tabelaDatasMax+="""
            <tr>
                <td align=center>{}</td>
                <td align=center>{}</td>
                <td align=center>{:.1f}</td>
            </tr>
        """.format(indMax+1, listaDatas[indicesMax[0][indMax]], maxCasos)
    
    tabelaDatasMin+="""
            </tbody>
        </table>
    """

    tabelaDatasMax+="""
            </tbody>
        </table>
    """
    return tabelaDatasMin, tabelaDatasMax

class PDF(FPDF, HTMLMixin):
    pass
    def header(self):
        dt = dataAtualF()
        self.image("header.png", x=0, y=0, alt_text="Cabecalho")
        self.write_html("""<br><br><br><br><br><br><br><br>
            <font color="black" size=35><p align=center>Relatório """+dt+"""</p></font>
            <br><br><br>""")

    def footer(self):
        self.set_y(-15)
        self.set_font("helvetica", "I", 10)
        self.cell(0, 10, f"Página {self.page_no()}/{{nb}} - Maurício Brito Teixeira", 0, 0, "C")

# ====Menu de relatorios====
def menuRelatorios(dicRelatorios):
    opc = 1
    while ( opc != 4 and opc > 0 and opc <6):
        print("\nGerenciamento de relatorios:\n")
        print("1 - Criar relatorio de hoje (Forma automatica)!")
        print("2 - Mostra um relatorio")
        print("3 - Mostra todos os relatorios")
        print("4 - Sair do menu de relatorios")
        print("5 - Teste")

        opc = int( input("Digite uma opção: ") )

        if opc ==1:
            BDconfiguracoes = recuperaConfiguracoes()
            recuperaUsuarios(BDusuariosRelatorio)
            confirma = input("Voce ja esta com o programa configurado? (S/N): ").upper()
            if confirma == 'N':
                print("Passo a passo para configurar!\n\n")
                print("1 - Leia o tutorial salvo na pasta desse projeto de nome 'tutorialConfiguracao.pdf'")
                print("2 - No menu principal de execucao va em 'Gerenciar configuracoes'")
                print("3 - Siga os passos explicados no pdf!!!")
            elif confirma == "S":
                faturamento = baixarArquivos()

                confirma = input("Deseja enviar email? (S/N): ").upper()

                if confirma == "S":
                    pyautogui.alert('O programa vai comecar, nao use nada do seu computador!')
                    pyautogui.press('win')
                    time.sleep(0.5)
                    linkGmail = "https://mail.google.com/"
                    pyperclip.copy(linkGmail)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(8)
                    pyautogui.click(int(BDconfiguracoes[6]), int(BDconfiguracoes[7]), clicks=1)
                    time.sleep(1)
                    for email in BDusuariosRelatorio:
                        pyautogui.write(email)
                        time.sleep(0.5)
                        pyautogui.press('tab')
                    time.sleep(0.5)
                    pyautogui.press('tab')
                    time.sleep(0.5)
                    pyautogui.write('Um faturamento ai')
                    time.sleep(0.5)
                    pyautogui.press('tab')
                    time.sleep(0.5)
                    pyautogui.write(f'o faturamento foi de: {faturamento}')
                    time.sleep(0.5)
                    pyautogui.press('tab')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    pyautogui.alert('Prontinho!')
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'w')
    
        elif opc == 2:
            data=input("Data a ser consultada: ")
            mostraRelatorio(dicRelatorios, data)
            pausa()
            
        elif opc == 3:
            mostraTodosRelatorios(dicRelatorios)
        
# =====================================================================

def criarRelatorioInterface():
    BDconfiguracoes = recuperaConfiguracoes()
    recuperaUsuarios(BDusuariosRelatorio)
    if len(BDconfiguracoes) > 9:
        baixarArquivos()
        
        tabela = pd.read_excel(r"C:\Users\Windows 10\Downloads\Dados-covid-19-estado.xlsx")
        tabela3 = pd.read_excel(r"C:\Users\Windows 10\Downloads\Dados-covid-19-municipios.xlsx")
        listaDatas = list(tabela["Data"])
        listaCasos = list(tabela["Casos por dia"])
        listaCasos = np.array(listaCasos)
        minCasos = tabela["Casos por dia"].min()
        maxCasos = tabela["Casos por dia"].max()
        indicesMin = np.where(listaCasos == int(minCasos))
        indicesMax = np.where(listaCasos == int(maxCasos))

        tabelaDatasMin, tabelaDatasMax = gerarTabelas(indicesMin, listaDatas, minCasos, indicesMax, maxCasos)
        grafico1(tabela)
        grafico2(tabela)
        grafico3(tabela)
        grafico4(tabela)
        grafico5(tabela, listaDatas)

        print("Porcentagem de casos minimos que tiveram comparado ao total: {:.2f}%".format(float((len(datasMinCasos)*100)/len(listaCasos))))

        totalE = tabela["Óbitos por dia"].sum()
        cidade = list(tabela3["Município"]).index("São Carlos")
        totalM = list(tabela3["Mun_Total de óbitos"])[cidade]
        print(totalE, totalM, list(tabela3["Município"])[cidade])

        print("Porcentagem de casos da cidade {} comparado ao total do estado: {:.2f}%".format(list(tabela3["Município"])[cidade], (totalM*100)/totalE))

        pdf = PDF()
        pdf.add_page()
        pdf.write_html(f"""<font color="black" size=25><p align=left>Gerais</p></font> <br>
                        <font size="12" color="#1a0dab"><p align=center><a href="{path}/GraficosRelatorios/casosPorDiaEstado-{dataAtualFArquivo()}.png">Casos por dia - Estado(Clique aqui)</a></p></font>
                        <font size="12" color="#1a0dab"><p align=center><a href="{path}/GraficosRelatorios/obitosPorDiaEstado-{dataAtualFArquivo()}.png">Óbitos por dia - Estado(Clique aqui)</a></p></font>
                        <font size="12" color="#1a0dab"><p align=center><a href="{path}/GraficosRelatorios/totalDeCasosEstado-{dataAtualFArquivo()}.png">Total de casos - Estado(Clique aqui)</a></p></font>""")
        pdf.image(path+"/GraficosRelatorios/dadosEstado-"+dataAtualFArquivo()+".png", h=100, w=200, x=10)
        pdf.write_html(tabelaDatasMin)
        pdf.write_html("<font color='green'><p align=center>{} dias de casos minimos equivale a {:.2f}% do total de dias em pandemia</p></font><br><br><br><br><br><br>".format(len(indicesMin[0]), float((len(datasMinCasos)*100)/len(listaCasos))))
        pdf.write_html(tabelaDatasMax)
        pdf.write_html("<font color='green'><p align=center>Isso equivale {:.2f}% do total de casos desde o inicio da pandemia!</p></font>".format(float((maxCasos)*100)/tabela["Casos por dia"].sum()))
        pdf.image(path+"/GraficosRelatorios/dadosComparativos-"+dataAtualFArquivo()+".png", h=100, w=180, x=10)
        if sum(dadosCausasMesRetrasado) > sum(dadosCausasMesPassado):
            pdf.write_html("<font color='green'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de casos referente ao mes retrasado!</p></font>".format(float(sum(dadosCausasMesRetrasado) - sum(dadosCausasMesPassado)), float((sum(dadosCausasMesPassado)*100)/sum(dadosCausasMesRetrasado))))
        else:
            pdf.write_html("<font color='red'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de casos referente ao mes retrasado!</p></font>".format(float(sum(dadosCausasMesPassado) - sum(dadosCausasMesRetrasado)), float((sum(dadosCausasMesRetrasado)*100)/sum(dadosCausasMesPassado))))
        if sum(dadosObitosMesRetrasado) > sum(dadosObitosMesPassado):
            pdf.write_html("<font color='green'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de óbitos referente ao mes retrasado!</p></font><br><br><br><br><br><br>".format(float(sum(dadosObitosMesRetrasado) - sum(dadosObitosMesPassado)), float((sum(dadosObitosMesPassado)*100)/sum(dadosObitosMesRetrasado))))
        else:
            pdf.write_html("<font color='red'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de óbitos referente ao mes retrasado!</p></font><br><br><br><br><br><br>".format(float(sum(dadosObitosMesPassado) - sum(dadosObitosMesRetrasado)), float((sum(dadosObitosMesRetrasado)*100)/sum(dadosObitosMesPassado))))
        pdf.output("relatorio.pdf")