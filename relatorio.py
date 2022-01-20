from datetime import datetime
from tkinter import Label
from tkinter.constants import W
from typing import Sized
from PIL.Image import new
import pyautogui
import time
import pyperclip
import pandas as pd

import os
from IPython.display import display
from pyscreeze import center
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
import shutil
import zipfile

import webbrowser
from pathlib import Path
import math

import requests
import json
from bs4 import BeautifulSoup

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
dadosCausasAntepenultimoMes = []
dadosObitosAntepenultimoMes = []
datasMaxCasos = []
tabelaDatasMin = """"""
tabelaDatasMax = """"""
diretorioDownloads = "C:\\Users\\Windows 10\\Downloads"
mesesEN = {'jan': (1), 'fev': (2), 'mar': (3), 'abr': (4), 'mai': (5), 'jun': (6), 'jul': (7), 'ago': (8), 'set': (9), 'out': (10), 'nov': (11), 'dez': (12)}
mesesNE = {'1': ('Janeiro'), '2': ('Fevereiro'), '3': ('Março'), '4': ('Abril'), '5': ('Maio'), '6': ('Junho'), '7': ('Julho'), '8': ('Agosto'), '9': ('Setembro'), '10': ('Outubro'), '11': ('Novembro'), '12': ('Dezembro')}
municipios = {} #Cod_IBGE - Mun_Total de casos - Mun_Total de óbitos
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
    if existeRelatorio(dic, data):
        alteraRelatorio(dic, data, dado1, "SIM")
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
        
def isnumber(value):
    try:
         float(value)
    except ValueError:
         return False
    return True

def baixarArquivos():
    BDconfiguracoes = recuperaConfiguracoes()
    recuperaUsuarios(BDusuariosRelatorio)
    pyautogui.PAUSE = 1

    pyautogui.alert('O programa vai comecar, nao use nada do seu computador!')

    pyautogui.press('win')
    time.sleep(1)
    link = "https://drive.google.com/drive/u/0/folders/1dSPGvocU1Tp-KobogpWfCBJF8QXEcSCY"
    pyperclip.copy(link)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.click(int(BDconfiguracoes[0]), int(BDconfiguracoes[1]), clicks=1)
    time.sleep(1) 
    pyautogui.click(int(BDconfiguracoes[2]), int(BDconfiguracoes[3]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[4]), int(BDconfiguracoes[5]), clicks=1)
    time.sleep(2)
    pyautogui.click(int(BDconfiguracoes[8]), int(BDconfiguracoes[9]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[2]), int(BDconfiguracoes[3]), clicks=1)
    time.sleep(1)
    pyautogui.click(int(BDconfiguracoes[4]), int(BDconfiguracoes[5]), clicks=1)
    time.sleep(2)
    # pyautogui.hotkey('ctrl', 'w')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'w')

def grafico1(tabela, dataRelatorioArquivo, datas):
    plt.plot(datas, tabela["Total de casos"], label="Total")
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"\\RelatoriosEgraficos\\totalDeCasosEstado-"+str(dataRelatorioArquivo)+".png")
    plt.close()

def grafico2(tabela, dataRelatorioArquivo, datas):
    plt.plot(datas, tabela["Casos por dia"], label="Casos")
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"\\RelatoriosEgraficos\\casosPorDiaEstado-"+str(dataRelatorioArquivo)+".png")
    plt.close()

def grafico3(tabela, dataRelatorioArquivo, datas):
    plt.plot(datas, tabela["Óbitos por dia"], label="Óbitos")
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"\\RelatoriosEgraficos\\obitosPorDiaEstado-"+str(dataRelatorioArquivo)+".png")
    plt.close()

def grafico4(tabela, dataRelatorioArquivo, datas):
    plt.plot(datas, tabela["Total de casos"], label="Total")
    plt.plot(datas, tabela["Casos por dia"], label="Casos")
    plt.plot(datas, tabela["Óbitos por dia"], label="Óbitos")
    plt.ylim(0, 30000)
    plt.xlabel("Data")
    plt.ylabel("Quantidade")
    plt.legend(loc="upper left")
    plt.savefig(path+"\\RelatoriosEgraficos\\dadosEstado-"+str(dataRelatorioArquivo)+".png")
    plt.close()

def grafico5(tabela, listaDatas, dataRelatorioArquivo):
    mesHoje = int(mesAtual())
    mesPassado = int(mesAtual())
    mesRetrasado = int(mesAtual())
    antepenultimoMes = int(mesAtual())
    if mesPassado == 3:
        mesPassado = 2
        mesRetrasado = 1
        antepenultimoMes = 12
    elif mesPassado == 2:
        mesPassado = 1
        mesRetrasado = 12
        antepenultimoMes = 11
    elif mesPassado == 1:
        mesPassado = 12
        mesRetrasado = 11
        antepenultimoMes = 10
    else: 
        mesPassado -= 1
        mesRetrasado -= 2
        antepenultimoMes -= 3
    continua = True
    x = int(len(listaDatas))-1
    while continua==True:
        dataFra = (str(listaDatas[x]).split('/'))
        if(isnumber(dataFra[1])==True):
            dataFra = int(dataFra[1])
        else:
            dataFra = int(mesesEN[str(dataFra[1])])
        if int(dataFra) != mesHoje:
            if int(dataFra) == mesPassado:
                datasMesPassado.append(tabela["Data"][x])
                dadosCausasMesPassado.append(int(tabela["Casos por dia"][x]))
                dadosObitosMesPassado.append(int(tabela["Óbitos por dia"][x]))
            if int(dataFra) == mesRetrasado:
                datasMesRetrasado.append(tabela["Data"][x])
                dadosCausasMesRetrasado.append(int(tabela["Casos por dia"][x]))
                dadosObitosMesRetrasado.append(int(tabela["Óbitos por dia"][x]))
            if int(dataFra) == antepenultimoMes:
                datasMesRetrasado.append(tabela["Data"][x])
                dadosCausasAntepenultimoMes.append(int(tabela["Casos por dia"][x]))
                dadosObitosAntepenultimoMes.append(int(tabela["Óbitos por dia"][x]))
            if x==0 or int(dataFra) < antepenultimoMes:
                continua = False
            else:
                x-=1
        else:
            x-=1
        if x==0:
            continua = False
        
    labels = ['Casos', 'Óbitos']
    antepenultimo = [sum(dadosCausasAntepenultimoMes), sum(dadosObitosAntepenultimoMes)]
    retrasado = [sum(dadosCausasMesRetrasado), sum(dadosObitosMesRetrasado)]
    passado = [sum(dadosCausasMesPassado), sum(dadosObitosMesPassado)]

    x = np.arange(len(labels))
    width = 0.25

    fig, ax = plt.subplots()
    rects1 = ax.bar(x + width/2, antepenultimo, width, label=f'{mesesNE[str(antepenultimoMes)]}')
    rects2 = ax.bar(x + (width*3)/2, retrasado, width, label=f'{mesesNE[str(mesRetrasado)]}')
    rects3 = ax.bar(x + (width*5)/2, passado, width, label=f'{mesesNE[str(mesPassado)]}')

    ax.set_ylabel('Quantidade')
    ax.set_title(f'Casos e óbitos - mes {mesesNE[str(antepenultimoMes)]}/{mesesNE[str(mesRetrasado)]}/{mesesNE[str(mesPassado)]}')
    ax.set_xticks(x+0.25+(0.25/2), labels)
    ax.legend()

    ax.bar_label(rects1, padding=3)
    ax.bar_label(rects2, padding=3)
    ax.bar_label(rects3, padding=3)

    fig.tight_layout()

    plt.savefig(path+"\\RelatoriosEgraficos\\dadosComparativos-"+str(dataRelatorioArquivo)+".png")
    plt.close()

def grafico6(legenda1, legenda2, qtd1, qtd2, titulo, nomeSalvar):
    labels = legenda1, legenda2
    sizes = [qtd1, qtd2]
    explode = (0, 0.1)

    fig1, ax1 = plt.subplots()
    ax1.set_title(titulo)
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')
    plt.savefig(path+"\\RelatoriosEgraficos\\"+str(nomeSalvar))
    plt.close()

def grafico7(legenda1, legenda2, legenda3, legenda4, qtd1, qtd2, qtd3, qtd4, titulo, nomeSalvar):
    print(legenda1, legenda2, legenda3, legenda4, qtd1, qtd2, qtd3, qtd4, titulo, nomeSalvar)
    labels = legenda1, legenda2, legenda3, legenda4
    sizes = [qtd1, qtd2, qtd3, qtd4]
    explode = (0, 0, 0, 0)

    fig1, ax1 = plt.subplots()
    ax1.set_title(titulo)
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')
    plt.savefig(path+"\\RelatoriosEgraficos\\"+str(nomeSalvar))
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

def enviarRelatorioEmail(dataParaArquivo, dataFormatada):
    recuperaUsuarios(BDusuariosRelatorio)
    diretorio = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}"
    os.makedirs(diretorio)
    casosAntes = f"{path}\\RelatoriosEgraficos\\casosPorDiaEstado-{dataParaArquivo}.png"
    casosDestino = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}\\casosPorDiaEstado-{dataParaArquivo}.png"
    obitosAntes = f"{path}\\RelatoriosEgraficos\\obitosPorDiaEstado-{dataParaArquivo}.png"
    obitosDestino = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}\\obitosPorDiaEstado-{dataParaArquivo}.png"
    totalDeCasosAntes = f"{path}\\RelatoriosEgraficos\\totalDeCasosEstado-{dataParaArquivo}.png"
    totalDeCasosDestino = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}\\totalDeCasosEstado-{dataParaArquivo}.png"
    dadosComparativosAntes = f"{path}\\RelatoriosEgraficos\\dadosComparativos-{dataParaArquivo}.png"
    dadosComparativosDestino = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}\\dadosComparativos-{dataParaArquivo}.png"
    dadosEstadoAntes = f"{path}\\RelatoriosEgraficos\\dadosEstado-{dataParaArquivo}.png"
    dadosEstadoDestino = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}\\dadosEstado-{dataParaArquivo}.png"
    relatorioAntes = f"{path}\\RelatoriosEgraficos\\relatorio-{dataParaArquivo}.pdf"
    relatorioDestino = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}\\relatorio-{dataParaArquivo}.pdf"
    diretorioZip = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}.zip"
    diretorioPastaZipar = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}"
    shutil.copy(casosAntes, casosDestino)
    shutil.copy(obitosAntes, obitosDestino)
    shutil.copy(totalDeCasosAntes, totalDeCasosDestino)
    shutil.copy(dadosComparativosAntes, dadosComparativosDestino)
    shutil.copy(dadosEstadoAntes, dadosEstadoDestino)
    shutil.copy(relatorioAntes, relatorioDestino)

    zip = zipfile.ZipFile(diretorioZip, 'w')
    for folder, subfolders, files in os.walk(diretorioPastaZipar):
        for file in files:
            zip.write(f"./RelatorioEnviar/relatorio-{str(dataParaArquivo)}/{file}", compress_type = zipfile.ZIP_DEFLATED)
    zip.close()
    pyautogui.alert('O programa vai comecar, nao use nada do seu computador!')
    pyautogui.press('win')
    time.sleep(2)
    linkGmail = "https://mail.google.com/"
    pyperclip.copy(linkGmail)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.click(int(BDconfiguracoes[6]), int(BDconfiguracoes[7]), clicks=1)
    time.sleep(1)
    for email in BDusuariosRelatorio:
        pyautogui.write(email)
        time.sleep(0.5)
        pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    sobre = f'Relatório do dia {dataFormatada}'
    pyperclip.copy(sobre)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    conteudo = f'Olá, boa tarde!'
    pyperclip.copy(conteudo)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(2.5)
    diretorioRelatorioZip = f"{path}\\RelatorioEnviar\\relatorio-{str(dataParaArquivo)}.zip"
    pyperclip.copy(diretorioRelatorioZip)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(2)
    pyautogui.hotkey('ctrl', 'w')
    shutil.rmtree(diretorioPastaZipar)
    os.remove(diretorioRelatorioZip)
    pyautogui.alert('Prontinho!')

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
class PDFMunicipios(FPDF, HTMLMixin):
    pass
    def header(self):
        dt = dataAtualF()
        self.image("header.png", x=0, y=0, alt_text="Cabecalho")
        self.write_html("""<br><br><br><br><br><br><br><br>
            <font color="black" size=35><p align=center>Municípios do Estado <br><br> de São Paulo</p></font>
            <br><br><br>""")

    def footer(self):
        self.set_y(-15)
        self.set_font("helvetica", "I", 10)
        self.cell(0, 10, f"Página {self.page_no()}/{{nb}} - Maurício Brito Teixeira", 0, 0, "C")

def gerarPdfMunicipios():
    diretorio = r"C:\Users\Windows 10\Downloads\Dados-covid-19-municipios.xlsx"
    if Path(diretorio).is_file():
        tabela3 = pd.read_excel(diretorio)

        tabela = """
            <table width="50%" border="1">
                <thead>
                    <tr>
                    <th width="20%">#</th>
                    <th width="80%">Município</th>
                    </tr>
                </thead>
                <tbody>
        """

        for indice in range(len(list(tabela3["Município"]))):
            tabela+="""
                <tr>
                    <td align=center>{}</td>
                    <td align=center>{}</td>
                </tr>
            """.format(indice+1, list(tabela3["Município"])[indice])
            municipios[list(tabela3["Município"])[indice]] = (list(tabela3["Cod_IBGE"])[indice], list(tabela3["Mun_Total de casos"])[indice], list(tabela3["Mun_Total de óbitos"])[indice])
        
        tabela+="""
                </tbody>
            </table>
        """

        pdf = PDFMunicipios()
        pdf.add_page()
        pdf.write_html(tabela)
        pdf.output(f"{path}\\Municipios\\Municipios de SP.pdf")
    else:
        print("O arquivo 'Dados-covid-19-municipios.xlsx' não existe!!!")

# ====Menu de relatorios====
def menuRelatorios(dicRelatorios):
    opc = 1
    while (opc != 6):
        print("\nGerenciamento de relatorios:\n")
        print("1 - Criar relatorio de hoje (Forma automatica)!")
        print("2 - Mostra um relatorio")
        print("3 - Mostra todos os relatorios")
        print("4 - Abrir")
        print("5 - Enviar")
        print("6 - Sair do menu de relatorios")

        opc = int( input("Digite uma opção: ") )

        if opc ==1:
            path = os.getcwd()
            dataRelatorioArquivo = dataAtualFArquivo()
            BDconfiguracoes = recuperaConfiguracoes()
            recuperaUsuarios(BDusuariosRelatorio)
            confirma = input("Voce ja esta com o programa configurado? (S/N): ").upper()
            if confirma == 'N':
                print("Passo a passo para configurar!\n\n")
                print("1 - Leia o tutorial salvo na pasta desse projeto de nome 'tutorialConfiguracao.pdf'")
                print("2 - No menu principal de execucao va em 'Gerenciar configuracoes'")
                print("3 - Siga os passos explicados no pdf!!!")
            elif confirma == "S":
                if len(BDconfiguracoes) > 9:
                    baixarArquivos()
                    tabelinha = pd.read_csv (f"{diretorioDownloads}\\Dados-covid-19-estado.csv", encoding="ANSI", sep=";")
                    tabelinha.to_excel(f"{diretorioDownloads}\\Dados-covid-19-estado.xlsx", index = None, header=True)
                    tabelinha3 = pd.read_csv (f"{diretorioDownloads}\\Dados-covid-19-municipios.csv", encoding="ANSI", sep=";")
                    tabelinha3.to_excel(f"{diretorioDownloads}\\Dados-covid-19-municipios.xlsx", index = None, header=True)

                    tabela = pd.read_excel(f"{diretorioDownloads}\\Dados-covid-19-estado.xlsx")
                    tabela3 = pd.read_excel(f"{diretorioDownloads}\\Dados-covid-19-municipios.xlsx")

                    gerarPdfMunicipios()

                    datas = []
                    ano = 0
                    mes = 0
                    dia = 0
                    for data in tabela["Data"]:
                        mes = (str(data).split('/'))
                        if(isnumber(mes[1])==True):
                            mes = int(mes[1])
                        else:
                            mes = int(mesesEN[str(mes[1])])
                        dia = int(str(data).split("/")[0])
                        ano = int(str(data).split("/")[2])
                        x = datetime(ano, mes, dia)
                        datas.append(x)

                    listaDatas = list(tabela["Data"])
                    listaCasos = list(tabela["Casos por dia"])
                    listaCasos = np.array(listaCasos)
                    minCasos = tabela["Casos por dia"].min()
                    maxCasos = tabela["Casos por dia"].max()
                    indicesMin = np.where(listaCasos == int(minCasos))
                    indicesMax = np.where(listaCasos == int(maxCasos))

                    tabelaDatasMin, tabelaDatasMax = gerarTabelas(indicesMin, listaDatas, minCasos, indicesMax, maxCasos)
                    grafico1(tabela, dataRelatorioArquivo, datas)
                    grafico2(tabela, dataRelatorioArquivo, datas)
                    grafico3(tabela, dataRelatorioArquivo, datas)
                    grafico4(tabela, dataRelatorioArquivo, datas)
                    grafico5(tabela, listaDatas, dataRelatorioArquivo)

                    totalEO = tabela["Óbitos por dia"].sum()
                    totalEC = tabela["Casos por dia"].sum()
                    cidade = list(tabela3["Município"]).index("São Carlos")
                    totalM = list(tabela3["Mun_Total de óbitos"])[cidade]

                    pdf = PDF()
                    pdf.add_page()
                    pdf.write_html(f"""<font color="black" size=30><p align=center>{totalEC} casos no estado de São Paulo</p></font> <br>""")
                    pdf.write_html(f"""<font color="black" size=30><p align=center>{totalEO} óbitos no estado de São Paulo</p></font> <br>""")
                    pdf.write_html(f"""<font color="black" size=25><p align=left>Gerais</p></font> <br>
                                    <font size="12" color="#1a0dab"><p align=center><a href="casosPorDiaEstado-{str(str(dataRelatorioArquivo))}.png">Casos por dia - Estado(Clique aqui)</a></p></font>
                                    <font size="12" color="#1a0dab"><p align=center><a href="obitosPorDiaEstado-{str(str(dataRelatorioArquivo))}.png">Óbitos por dia - Estado(Clique aqui)</a></p></font>
                                    <font size="12" color="#1a0dab"><p align=center><a href="totalDeCasosEstado-{str(str(dataRelatorioArquivo))}.png">Total de casos - Estado(Clique aqui)</a></p></font>""")
                    pdf.image(path+"/RelatoriosEgraficos/dadosEstado-"+str(dataRelatorioArquivo)+".png", h=100, w=200, x=10)
                    pdf.write_html(tabelaDatasMin)
                    pdf.write_html("<font color='green'><p align=center>{} dias de casos minimos equivale a {:.2f}% do total de dias em pandemia</p></font><br><br><br><br><br><br>".format(len(indicesMin[0]), float((len(datasMinCasos)*100)/len(listaCasos))))
                    pdf.write_html(tabelaDatasMax)
                    pdf.write_html("<font color='green'><p align=center>Isso equivale {:.2f}% do total de casos desde o inicio da pandemia!</p></font>".format(float((maxCasos)*100)/tabela["Casos por dia"].sum()))
                    pdf.image(path+"/RelatoriosEgraficos/dadosComparativos-"+str(dataRelatorioArquivo)+".png", h=100, w=180, x=10)
                    if sum(dadosCausasMesRetrasado) > sum(dadosCausasMesPassado):
                        pdf.write_html("<font color='green'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de casos referente ao mes retrasado!</p></font>".format(float(sum(dadosCausasMesRetrasado) - sum(dadosCausasMesPassado)), float(float(sum(dadosCausasMesRetrasado) - sum(dadosCausasMesPassado))*100)/sum(dadosCausasMesRetrasado)))
                    else:
                        pdf.write_html("<font color='red'><p align=center>Mes passado teve um aumento de {:.1f}({:.2f}%) de casos referente ao mes retrasado!</p></font>".format(float(sum(dadosCausasMesPassado) - sum(dadosCausasMesRetrasado)), float(float(sum(dadosCausasMesPassado) - sum(dadosCausasMesRetrasado))*100)/sum(dadosCausasMesPassado)))
                    if sum(dadosObitosMesRetrasado) > sum(dadosObitosMesPassado):
                        pdf.write_html("<font color='green'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de óbitos referente ao mes retrasado!</p></font><br><br><br><br><br><br>".format(float(sum(dadosObitosMesRetrasado) - sum(dadosObitosMesPassado)), float(float(sum(dadosObitosMesRetrasado) - sum(dadosObitosMesPassado))*100)/sum(dadosObitosMesRetrasado)))
                    else:
                        pdf.write_html("<font color='red'><p align=center>Mes passado teve um aumento de {:.1f}({:.2f}%) de óbitos referente ao mes retrasado!</p></font><br><br><br><br><br><br>".format(float(sum(dadosObitosMesPassado) - sum(dadosObitosMesRetrasado)), float(float(sum(dadosObitosMesPassado) - sum(dadosObitosMesRetrasado))*100)/sum(dadosObitosMesPassado)))
                    pdf.image("linha.png", h=50, w=180, x=10)

                    url = requests.get('https://ibge.gov.br/explica/desemprego.php')
                    html = url.content
                    site = BeautifulSoup(html, 'html.parser')
                    text = str(site).find('var pizzaData = ')
                    dadosPizza = str(str(site)[(text + 17):str(site)[text:].find('];') + text]).split('{')
                    totalDados = 0
                    labels = []
                    qtds = []

                    for a in range(1, len(dadosPizza)):
                        dadosJson = (json.loads("{"+str(dadosPizza[a]).replace('},', '}')))
                        totalDados += int(dadosJson["numPessoas"])
                        labels.append(dadosJson["status"])
                        qtds.append(int(dadosJson["numPessoas"]))

                    grafico7(labels[0], labels[1], labels[2], labels[3], int(qtds[0]), int(qtds[1]), int(qtds[2]), int(qtds[3]), "População brasileira, de acordo com as divisões do mercado de trabalho, 3º trimestre 2021", f"desemprego-IBGE-{dataRelatorioArquivo}.png")

                    pdf.image(path+"/RelatoriosEgraficos/desemprego-IBGE-"+dataRelatorioArquivo+".png", h=100, w=180, x=10)

                    texto = "São varias as consequencias causadas pela pandemia em todo o mundo, O que é notório é que seus efeitos trágicos impactaram na vida social, econômica e profissional das pessoas em todo mundo. \n\nA economia mundial terá enormes perdas, pois as atividades agrícolas, comerciais, industriais e de turismo estão sofrendo uma queda bem grande na sua produtividade por conta do isolamento, as ações da Bolsa de Valores oscilarão grandes variações.\n\nFatores que poderão causar em um futuro próximo, principalmente nas nações menos desenvolvidas economicamente, desemprego, aumento dos índices inflacionários, escassez de alimentos e crescimento da desigualdade social, da violência urbana e criminalidade. \n\nDesemprego por exemplo, mesmo antes de pandemia no Brasil, o desemprego já era um problema na vida de boa parte da população. Afinal, a economia não estava indo bem há um tempo, e como podemos ver nos graficos, com o alto indice de casos do novo coronavírus, muitas vagas de emprego foram fechadas. Por conta disso, só nos primeiros meses do ano de 2020, segundo o IBGE, quase 5 milhões de pessoas tiveram que parar de trabalhar \n\nA pandemia no Brasil fez a desigualdade se tornar ainda mais evidente. Isso porque ficou claro que nem todo mundo tem condições de seguir as mesmas medidas de prevenção. seguimos com mais alguns dados no relatorio"
                    pdf.multi_cell(w=140, h=8, txt=texto, align='J')

                    pdf.add_page()

                    pdf.write_html(f"""<font color="black" size=25><p align=left>Municipios de usuarios</p></font> <br>""")

                    for user in BDusuariosRelatorio.keys():
                        if BDusuariosRelatorio[user][1] in municipios.keys():
                            url = host+"/news/search"

                            querystring = {"q":f"Covid 19 "+BDusuariosRelatorio[user][1],"freshness":"Day","textFormat":"Raw","safeSearch":"Off"}

                            headers = {
                                'x-bingapis-sdk': "true",
                                'x-rapidapi-host': str(hostCovid19),
                                'x-rapidapi-key': str(chaveCovid19)
                                }

                            response = requests.request("GET", url, headers=headers, params=querystring)

                            pdf.write_html(f"""<font color="black" size=35><p align=center>{BDusuariosRelatorio[user][1]}</p></font> <br>""")
                            totalDeCasos = municipios[BDusuariosRelatorio[user][1]][1]
                            totalDeObitos = municipios[BDusuariosRelatorio[user][1]][2]
                            ptotalM = (totalDeCasos*100)/totalEC
                            ptotalE = ((totalEC - totalDeCasos)*100)/totalEC
                            ptotaMC = (totalDeObitos*100)/totalDeCasos
                            grafico6("Resto do Estado", BDusuariosRelatorio[user][1], ptotalE, ptotalM, f"Casos de {BDusuariosRelatorio[user][1]} comparado ao resto do estado", f"casos-{BDusuariosRelatorio[user][1]}-estado-{dataRelatorioArquivo}.png")
                            grafico6("Casos", "Mortes", 100-ptotaMC, ptotaMC, f"Casos/Óbitos - {totalDeCasos}/{totalDeObitos}", f"obitos-casos-{BDusuariosRelatorio[user][1]}-{dataRelatorioArquivo}.png")
                            pdf.image(path+"/RelatoriosEgraficos/"+f"casos-{BDusuariosRelatorio[user][1]}-estado-{dataRelatorioArquivo}.png", h=100, w=180, x=10)
                            pdf.write_html("""<font color="black" size=10><p align=center>{} equivale à {:.2f}% dos casos do estado de São Paulo({})</p></font> <br>""".format(totalDeCasos, ptotalM, totalEC))
                            pdf.image(path+"/RelatoriosEgraficos/"+f"obitos-casos-{BDusuariosRelatorio[user][1]}-{dataRelatorioArquivo}.png", h=100, w=180, x=10)
                            pdf.write_html(f"""<font color="black" size=10><p align=center>A cada {math.ceil(totalDeCasos/totalDeObitos)} casos de Covid-19 1 morre! (Aproximadamente)</p></font> <br>""")

                            pdf.write_html(f"""<font color="green" size=20><p align=center>Notícias do Covid na cidade de {BDusuariosRelatorio[user][1]}</p></font> <br>""")
                            
                            for a in range(0, len(json.loads(response.text)["value"])):
                                try:
                                    titulo = json.loads(response.text)["value"][a]["name"]
                                    url = json.loads(response.text)["value"][a]["url"]
                                    fonte = json.loads(response.text)["value"][a]["provider"][0]["name"]
                                    description = json.loads(response.text)["value"][a]["description"]
                                    imagem = json.loads(response.text)["value"][a]["image"]["thumbnail"]["contentUrl"]
                                    pdf.image(imagem, h=35, w=35, x=85)
                                    pdf.write_html(f"""
                                    <ul><li>{fonte}</li>
                                        <ol>
                                            <li>{titulo}</li>
                                            <li>{description}</li>
                                            <li><a href='{url}'>Ver notícia</a></li>
                                        </ol>
                                    </ul>
                                    <br><br><br>
                                    """)
                                except:
                                    print("erro!")
                            pdf.add_page()
                        else:
                            print(f"O municipio '{BDusuariosRelatorio[user][1]}' do usuario de email '{user}' é invalido!")
                    pdf.set_font('times', '', 25)
                    pdf.multi_cell(w=190, h=25, txt="Vida social, econômica e profissional das pessoas vão ter grandes impactos gerados pela pandemia no mundo todo!", align='C')
                    diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{dataRelatorioArquivo}.pdf"
                    pdf.output(diretorio)

                    confirma = input("Deseja enviar email? (S/N): ").upper()

                    if confirma == "S":
                        enviarRelatorioEmail(dataRelatorioArquivo, dataAtualF())
                        insereRelatorio(dicRelatorios, "SIM", "NÃO")
                    else:
                        insereRelatorio(dicRelatorios, "NÃO", "NÃO")

        elif opc == 2:
            data=input("Data a ser consultada: ")
            mostraRelatorio(dicRelatorios, data)
            pausa()
            
        elif opc == 3:
            mostraTodosRelatorios(dicRelatorios)

        elif opc == 4:
            data = input("Informe a data do relatorio no formato 'dd/mm/yyyy': ")
            data = str(data).replace('/', '-')
            path = os.getcwd()
            diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{data}.pdf"
            if Path(diretorio).is_file():
                webbrowser.open(diretorio)
            else:
                print("O arquivo não existe!!!")
            
        elif opc == 5:
            data = input("Informe a data do relatorio no formato 'dd/mm/yyyy': ")
            dataArq = str(data).replace('/', '-')
            path = os.getcwd()
            diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{dataArq}.pdf"
            if Path(diretorio).is_file():
                enviarRelatorioEmail(dataArq, data)
                respData, respEnviou, respAtualizou = mostraRelatorioInterface(dicRelatorios, data)
                alteraRelatorioInterface(dicRelatorios, respData, "SIM", respAtualizou)
                gravaRelatorios(dicRelatorios)
                recuperaRelatorios(dicRelatorios)
                print("Relatorio enviado com sucesso")
            else:
                print("O arquivo não existe!!!")

        elif opc == 6:
            gravaRelatorios(dicRelatorios)
        
# =====================================================================

def criarRelatorioInterface():
    dataRelatorioArquivo = dataAtualFArquivo()
    BDconfiguracoes = recuperaConfiguracoes()
    recuperaUsuarios(BDusuariosRelatorio)
    if len(BDconfiguracoes) > 9:
        baixarArquivos()
        tabelinha = pd.read_csv (f"{diretorioDownloads}\\Dados-covid-19-estado.csv", encoding="ANSI", sep=";")
        tabelinha.to_excel(f"{diretorioDownloads}\\Dados-covid-19-estado.xlsx", index = None, header=True)
        tabelinha3 = pd.read_csv (f"{diretorioDownloads}\\Dados-covid-19-municipios.csv", encoding="ANSI", sep=";")
        tabelinha3.to_excel(f"{diretorioDownloads}\\Dados-covid-19-municipios.xlsx", index = None, header=True)

        tabela = pd.read_excel(f"{diretorioDownloads}\\Dados-covid-19-estado.xlsx")
        tabela3 = pd.read_excel(f"{diretorioDownloads}\\Dados-covid-19-municipios.xlsx")

        gerarPdfMunicipios()

        datas = []
        ano = 0
        mes = 0
        dia = 0
        for data in tabela["Data"]:
            mes = (str(data).split('/'))
            if(isnumber(mes[1])==True):
                mes = int(mes[1])
            else:
                mes = int(mesesEN[str(mes[1])])
            dia = int(str(data).split("/")[0])
            ano = int(str(data).split("/")[2])
            x = datetime(ano, mes, dia)
            datas.append(x)

        listaDatas = list(tabela["Data"])
        listaCasos = list(tabela["Casos por dia"])
        listaCasos = np.array(listaCasos)
        minCasos = tabela["Casos por dia"].min()
        maxCasos = tabela["Casos por dia"].max()
        indicesMin = np.where(listaCasos == int(minCasos))
        indicesMax = np.where(listaCasos == int(maxCasos))

        tabelaDatasMin, tabelaDatasMax = gerarTabelas(indicesMin, listaDatas, minCasos, indicesMax, maxCasos)
        grafico1(tabela, dataRelatorioArquivo, datas)
        grafico2(tabela, dataRelatorioArquivo, datas)
        grafico3(tabela, dataRelatorioArquivo, datas)
        grafico4(tabela, dataRelatorioArquivo, datas)
        grafico5(tabela, listaDatas, dataRelatorioArquivo)

        totalEO = tabela["Óbitos por dia"].sum()
        totalEC = tabela["Casos por dia"].sum()
        cidade = list(tabela3["Município"]).index("São Carlos")
        totalM = list(tabela3["Mun_Total de óbitos"])[cidade]

        pdf = PDF()
        pdf.add_page()
        pdf.write_html("""<font color="black" size=25><p align=left>Gerais</p></font> <br>""")
        pdf.write_html(f"""<font color="black" size=20><p align=center>{totalEC} casos no estado de São Paulo</p></font> <br>""")
        pdf.write_html(f"""<font color="black" size=20><p align=center>{totalEO} óbitos no estado de São Paulo</p></font> <br>""")
        pdf.write_html(f"""<font size="12" color="#1a0dab"><p align=center><a href="casosPorDiaEstado-{str(str(dataRelatorioArquivo))}.png">Casos por dia - Estado(Clique aqui)</a></p></font>
                        <font size="12" color="#1a0dab"><p align=center><a href="obitosPorDiaEstado-{str(str(dataRelatorioArquivo))}.png">Óbitos por dia - Estado(Clique aqui)</a></p></font>
                        <font size="12" color="#1a0dab"><p align=center><a href="totalDeCasosEstado-{str(str(dataRelatorioArquivo))}.png">Total de casos - Estado(Clique aqui)</a></p></font>""")
        pdf.image(path+"/RelatoriosEgraficos/dadosEstado-"+str(dataRelatorioArquivo)+".png", h=100, w=200, x=10)
        pdf.write_html(tabelaDatasMin)
        pdf.write_html("<font color='green'><p align=center>{} dias de casos minimos equivale a {:.2f}% do total de dias em pandemia</p></font><br><br><br><br><br><br>".format(len(indicesMin[0]), float((len(datasMinCasos)*100)/len(listaCasos))))
        pdf.write_html(tabelaDatasMax)
        pdf.write_html("<font color='green'><p align=center>Isso equivale {:.2f}% do total de casos desde o inicio da pandemia!</p></font>".format(float((maxCasos)*100)/tabela["Casos por dia"].sum()))
        pdf.image(path+"/RelatoriosEgraficos/dadosComparativos-"+str(dataRelatorioArquivo)+".png", h=100, w=180, x=10)
        if sum(dadosCausasMesRetrasado) > sum(dadosCausasMesPassado):
            pdf.write_html("<font color='green'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de casos referente ao mes retrasado!</p></font>".format(float(sum(dadosCausasMesRetrasado) - sum(dadosCausasMesPassado)), float(float(sum(dadosCausasMesRetrasado) - sum(dadosCausasMesPassado))*100)/sum(dadosCausasMesRetrasado)))
        else:
            pdf.write_html("<font color='red'><p align=center>Mes passado teve um aumento de {:.1f}({:.2f}%) de casos referente ao mes retrasado!</p></font>".format(float(sum(dadosCausasMesPassado) - sum(dadosCausasMesRetrasado)), float(float(sum(dadosCausasMesPassado) - sum(dadosCausasMesRetrasado))*100)/sum(dadosCausasMesPassado)))
        if sum(dadosObitosMesRetrasado) > sum(dadosObitosMesPassado):
            pdf.write_html("<font color='green'><p align=center>Mes passado teve uma diminuicao de {:.1f}({:.2f}%) de óbitos referente ao mes retrasado!</p></font><br><br><br><br><br><br>".format(float(sum(dadosObitosMesRetrasado) - sum(dadosObitosMesPassado)), float(float(sum(dadosObitosMesRetrasado) - sum(dadosObitosMesPassado))*100)/sum(dadosObitosMesRetrasado)))
        else:
            pdf.write_html("<font color='red'><p align=center>Mes passado teve um aumento de {:.1f}({:.2f}%) de óbitos referente ao mes retrasado!</p></font><br><br><br><br><br><br>".format(float(sum(dadosObitosMesPassado) - sum(dadosObitosMesRetrasado)), float(float(sum(dadosObitosMesPassado) - sum(dadosObitosMesRetrasado))*100)/sum(dadosObitosMesPassado)))
        pdf.image("linha.png", h=50, w=180, x=10)
         
        url = requests.get('https://ibge.gov.br/explica/desemprego.php')
        html = url.content
        site = BeautifulSoup(html, 'html.parser')
        text = str(site).find('var pizzaData = ')
        dadosPizza = str(str(site)[(text + 17):str(site)[text:].find('];') + text]).split('{')
        totalDados = 0
        labels = []
        qtds = []

        for a in range(1, len(dadosPizza)):
            dadosJson = (json.loads("{"+str(dadosPizza[a]).replace('},', '}')))
            totalDados += int(dadosJson["numPessoas"])
            labels.append(dadosJson["status"])
            qtds.append(int(dadosJson["numPessoas"]))

        grafico7(labels[0], labels[1], labels[2], labels[3], int(qtds[0]), int(qtds[1]), int(qtds[2]), int(qtds[3]), "População brasileira, de acordo com as divisões do mercado de trabalho, 3º trimestre 2021", f"desemprego-IBGE-{dataRelatorioArquivo}.png")

        pdf.add_page()

        pdf.image(path+"/RelatoriosEgraficos/desemprego-IBGE-"+dataRelatorioArquivo+".png", h=100, w=180, x=10)

        texto = "São varias as consequencias causadas pela pandemia em todo o mundo, O que é notório é que seus efeitos trágicos impactaram na vida social, econômica e profissional das pessoas em todo mundo. \n\nA economia mundial terá enormes perdas, pois as atividades agrícolas, comerciais, industriais e de turismo estão sofrendo uma queda bem grande na sua produtividade por conta do isolamento, as ações da Bolsa de Valores oscilarão grandes variações.\n\nFatores que poderão causar em um futuro próximo, principalmente nas nações menos desenvolvidas economicamente, desemprego, aumento dos índices inflacionários, escassez de alimentos e crescimento da desigualdade social, da violência urbana e criminalidade. \n\nDesemprego por exemplo, mesmo antes de pandemia no Brasil, o desemprego já era um problema na vida de boa parte da população. Afinal, a economia não estava indo bem há um tempo, e como podemos ver nos graficos, com o alto indice de casos do novo coronavírus, muitas vagas de emprego foram fechadas. Por conta disso, só nos primeiros meses do ano de 2020, segundo o IBGE, quase 5 milhões de pessoas tiveram que parar de trabalhar \n\nA pandemia no Brasil fez a desigualdade se tornar ainda mais evidente. Isso porque ficou claro que nem todo mundo tem condições de seguir as mesmas medidas de prevenção. seguimos com mais alguns dados no relatorio"
        pdf.multi_cell(w=0, h=8, txt=texto, align='J')

        pdf.add_page()

        pdf.write_html(f"""<font color="black" size=25><p align=left>Municipios de usuarios</p></font> <br>""")
        for user in BDusuariosRelatorio.keys():
            if BDusuariosRelatorio[user][1] in municipios.keys():
                url = host+"/news/search"

                querystring = {"q":f"Covid 19 "+BDusuariosRelatorio[user][1],"freshness":"Day","textFormat":"Raw","safeSearch":"Off"}

                headers = {
                    'x-bingapis-sdk': "true",
                    'x-rapidapi-host': str(hostCovid19),
                    'x-rapidapi-key': str(chaveCovid19)
                    }

                response = requests.request("GET", url, headers=headers, params=querystring)

                pdf.write_html(f"""<font color="black" size=35><p align=center>{BDusuariosRelatorio[user][1]}</p></font> <br>""")
                totalDeCasos = municipios[BDusuariosRelatorio[user][1]][1]
                totalDeObitos = municipios[BDusuariosRelatorio[user][1]][2]
                ptotalM = (totalDeCasos*100)/totalEC
                ptotalE = ((totalEC - totalDeCasos)*100)/totalEC
                ptotaMC = (totalDeObitos*100)/totalDeCasos
                grafico6("Resto do Estado", BDusuariosRelatorio[user][1], ptotalE, ptotalM, f"Casos de {BDusuariosRelatorio[user][1]} comparado ao resto do estado", f"casos-{BDusuariosRelatorio[user][1]}-estado-{dataRelatorioArquivo}.png")
                grafico6("Casos", "Mortes", 100-ptotaMC, ptotaMC, f"Casos/Óbitos - {totalDeCasos}/{totalDeObitos}", f"obitos-casos-{BDusuariosRelatorio[user][1]}-{dataRelatorioArquivo}.png")
                pdf.image(path+"/RelatoriosEgraficos/"+f"casos-{BDusuariosRelatorio[user][1]}-estado-{dataRelatorioArquivo}.png", h=100, w=180, x=10)
                pdf.write_html("""<font color="black" size=10><p align=center>{} equivale à {:.2f}% dos casos do estado de São Paulo({})</p></font> <br>""".format(totalDeCasos, ptotalM, totalEC))
                pdf.image(path+"/RelatoriosEgraficos/"+f"obitos-casos-{BDusuariosRelatorio[user][1]}-{dataRelatorioArquivo}.png", h=100, w=180, x=10)
                pdf.write_html(f"""<font color="black" size=10><p align=center>A cada {math.ceil(totalDeCasos/totalDeObitos)} casos de Covid-19 1 morre! (Aproximadamente)</p></font> <br>""")

                pdf.write_html(f"""<font color="green" size=20><p align=center>Notícias do Covid na cidade de {BDusuariosRelatorio[user][1]}</p></font> <br>""")
                
                for a in range(0, len(json.loads(response.text)["value"])):
                    try:
                        titulo = json.loads(response.text)["value"][a]["name"]
                        url = json.loads(response.text)["value"][a]["url"]
                        fonte = json.loads(response.text)["value"][a]["provider"][0]["name"]
                        description = json.loads(response.text)["value"][a]["description"]
                        imagem = json.loads(response.text)["value"][a]["image"]["thumbnail"]["contentUrl"]
                        pdf.image(imagem, h=35, w=35, x=85)
                        pdf.write_html(f"""
                        <ul><li>{fonte}</li>
                            <ol>
                                <li>{titulo}</li>
                                <li>{description}</li>
                                <li><a href='{url}'>Ver notícia</a></li>
                            </ol>
                        </ul>
                        <br><br><br>
                        """)
                    except:
                        print("erro!")
                    
                pdf.add_page()
                
            else:
                print(f"O municipio '{BDusuariosRelatorio[user][1]}' do usuario de email '{user}' é invalido!")
        
        pdf.set_font('times', '', 25)
        pdf.multi_cell(w=190, h=8, txt="Vida social, econômica e profissional das pessoas vão ter grandes impactos gerados pela pandemia no mundo todo!", align='C')
        diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{dataRelatorioArquivo}.pdf"
        pdf.output(diretorio)
        pyautogui.alert('Volte à tela do programa!')
        okcancel = messagebox.askokcancel("Atenção", "Deseja enviar email?")
        if okcancel == True:
            enviarRelatorioEmail(dataRelatorioArquivo, dataAtualF())
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-estado.csv")
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-estado.xlsx")
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-municipios.csv")
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-municipios.xlsx")
            return dataAtualF(), "SIM", "NÃO"
        else:
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-estado.csv")
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-estado.xlsx")
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-municipios.csv")
            os.remove(f"{diretorioDownloads}\\Dados-covid-19-municipios.xlsx")
            return dataAtualF(), "NÃO", "NÃO"
        

# ====ALTERA RELATORIO INTERFACE====
def alteraRelatorioInterface(dic,chave, dado1, dado2):
    if existeRelatorio(dic,chave):
        dic[chave]=(dado1, dado2)
        print("Dados alterados com sucesso!")
    else:
        print("Relatorio não cadastrado!")

# =======================INSERE RELATORIO=========================
def insereRelatorioInterface(dic, data, dado1, dado2):
    if existeRelatorio(dic, data):
        alteraRelatorioInterface(dic,data, dado1, "SIM")
    else:
        dic[data]=(dado1, dado2)
        print("Dados inseridos com sucesso!")
    
# ====REMOVE UM RELATORIO INTERFACE====
def removeRelatorioInterface(dic,chave):
    if existeRelatorio(dic,chave):
        okcancel = messagebox.askokcancel("Atenção", "Tem certeza que deseja excluir?")
        if okcancel == True:
            del dic[chave]
            print("Dados apagados com sucesso!")
        else:
            print("Exclusão cancelada!")
    else:
        print("Relatorio não cadastrado!")

# ====EXIBE UM RELATORIO INTERFACE====
def mostraRelatorioInterface(dic,chave):
    if existeRelatorio(dic,chave):
        dados = dic[chave]
        return chave, dados[0], dados[1]
    else:
        print("Relatorio não cadastrada!")
        return False, False, False