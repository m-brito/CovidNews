import pyautogui
import time
import pyperclip
import pandas as pd

import os
from IPython.display import display

from gerais import *
from configuracoes import *
from usuario import *

BDconfiguracoes = recuperaConfiguracoes()
BDusuariosRelatorio = {}
recuperaUsuarios(BDusuariosRelatorio)

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

# ====Menu de relatorios====
def menuRelatorios(dicRelatorios):
    opc = 1
    while ( opc != 4 and opc > 0 and opc <4):
        print("\nGerenciamento de relatorios:\n")
        print("1 - Criar relatorio de hoje (Forma automatica)!")
        print("2 - Mostra um relatorio")
        print("3 - Mostra todos os relatorios")
        print("4 - Sair do menu de relatorios")

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
                # pyautogui.hotkey('ctrl', 'w')
                time.sleep(2)
                pyautogui.alert('O programa foi executado, voce ja pode usar normalmente!')

                tabela = pd.read_excel(r"C:\Users\Windows 10\Downloads\Vendas - Dez.xlsx")
                display(tabela)
                faturamento = tabela["Valor Final"].sum()

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