import pyautogui
import time
import pyperclip
import pandas as pd

import os
from IPython.display import display
from gerais import *
from tkinter import messagebox


# ===================Configuracoes==============
# ====Insere configuracoes====
def insereConfiguracoes(listaConfiguracoes, positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y):
    if len(listaConfiguracoes) > 0:
        print("Configuracao já cadastrada!")
        pausa()
    else:
        listaConfiguracoes = [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]
        return listaConfiguracoes

# ====Altera configuracoes====
def alteraConfiguracoes(listaConfiguracoes, positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y):
    if len(listaConfiguracoes) > 0:
        mostraTodasConfiguracoes(listaConfiguracoes)
        confirma = input("Tem certeza que deseja atualiza-la? (S/N): ").upper()
        if confirma == 'S':
            listaConfiguracoes = [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]
            return listaConfiguracoes
        else:
            print("Alteração cancelada!")
            pausa()
    else:
        print("Configuracoes não cadastradas!")
        pausa()

# ====Remove configuracoes====
def removeConfiguracoes(listaConfiguracoes):
    if len(listaConfiguracoes) > 0:
        listaConfiguracoes = []
        return listaConfiguracoes
    else:
        print("Configuracoes não cadastradas!")
        pausa()

# ====Mostra todas configuracoes====
def mostraTodasConfiguracoes(listConfiguracoes):
    if len(listConfiguracoes)>0:
        print("Relatório: Todas as configuracoes\n")
        print(f"Posicao do arquivo 1 X: {listConfiguracoes[0]}")
        print(f"Posicao do arquivo 1 Y: {listConfiguracoes[1]}")
        print(f"Posicao do arquivo 2 X: {listConfiguracoes[8]}")
        print(f"Posicao do arquivo 2 Y: {listConfiguracoes[9]}")
        print(f"Posicao dos tres pontos X: {listConfiguracoes[2]}")
        print(f"Posicao dos tres pontos Y: {listConfiguracoes[3]}")
        print(f"Posicao do 'fazer download' X: {listConfiguracoes[4]}")
        print(f"Posicao do 'fazer download' Y: {listConfiguracoes[5]}")
        print(f"Posicao do 'Criar email' X: {listConfiguracoes[6]}")
        print(f"Posicao do 'Criar email' Y: {listConfiguracoes[7]}")
        print("")
        pausa()
    else:
        print("Nada configurado ainda")

# ======Grava dados no arquivo=====
def gravaConfiguracoes(listaConfiguracoes):
    if len(listaConfiguracoes) > 0:
        arq = open("configuracoes.txt", "w")
        linha = f"{listaConfiguracoes[0]};{listaConfiguracoes[1]};{listaConfiguracoes[2]};{listaConfiguracoes[3]};{listaConfiguracoes[4]};{listaConfiguracoes[5]};{listaConfiguracoes[6]};{listaConfiguracoes[7]};{listaConfiguracoes[8]};{listaConfiguracoes[9]}\n"
        arq.write(linha)
        arq.close()
    else:
        arq = open("configuracoes.txt", "w")
        linha = ""
        arq.write(linha)
        arq.close()

# ======Pega dados do arquivo====
def recuperaConfiguracoes():
    if (existe_arquivo("configuracoes.txt")):
        arq = open("configuracoes.txt", "r")
        listaConfiguracoes = []
        for linha in arq:
            linha = linha[:len(linha)-1]
            lista = linha.split(";")
            positionArqX = lista[0]
            positionArqY = lista[1]
            positionArq2X = lista[8]
            positionArq2Y = lista[9]
            positionTresPontosX = lista[2]
            positionTresPontosY = lista[3]
            positionFazerDownloadX = lista[4]
            positionFazerDownloadY = lista[5]
            positionEmailX = lista[6]
            positionEmailY = lista[7]
            listaConfiguracoes = [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]
        return listaConfiguracoes
    else:
        return []

def retornarConfigsParaPrograma():
    pyautogui.PAUSE = 1

    pyautogui.alert('Siga os passos corretamentes para não ter erros na configuracao!')
    pyautogui.alert('Espere ate a mensagem de concluido para tirar o mouse do local correto!')
    pyautogui.alert('O programa vai comecar, não mexa nem no mouse nem no teclado a menos que apareca uma mensagem pedindo!')

    pyautogui.press('win')
    time.sleep(0.5)
    link = "https://drive.google.com/drive/u/0/folders/1dSPGvocU1Tp-KobogpWfCBJF8QXEcSCY"
    pyperclip.copy(link)
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')

    pyautogui.alert('Coloque o mouse no arquivo 1!')

    time.sleep(5)

    pontoMouse = pyautogui.position()
    positionArqX = list(pontoMouse)[0]
    positionArqY = list(pontoMouse)[1]

    time.sleep(2)

    pyautogui.alert('Coloque o mouse no arquivo 2!')

    time.sleep(5)

    pontoMouse = pyautogui.position()
    positionArq2X = list(pontoMouse)[0]
    positionArq2Y = list(pontoMouse)[1]

    pyautogui.alert('Clique no arquivo 1!')
    time.sleep(2)
    pyautogui.alert('Coloque o mouse nos tres pontinhos!')

    time.sleep(5)

    pontoMouse = pyautogui.position()

    # pyautogui.alert(pontoMouse)
    
    positionTresPontosX = list(pontoMouse)[0]
    positionTresPontosY = list(pontoMouse)[1]

    pyautogui.alert('Clique nos tres pontinhos e coloque o mouse no "fazer download"!')

    time.sleep(5)

    pontoMouse = pyautogui.position()

    # pyautogui.alert(pontoMouse)

    positionFazerDownloadX = list(pontoMouse)[0]
    positionFazerDownloadY = list(pontoMouse)[1]

    pyautogui.hotkey('ctrl', 't')
    time.sleep(0.5)
    linkGmail = "https://mail.google.com/"
    pyperclip.copy(linkGmail)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(8)
    pyautogui.alert('Coloque o mouse no "Escrever"!')
    time.sleep(5)
    pontoMouse = pyautogui.position()
    positionEmailX = list(pontoMouse)[0]
    positionEmailY = list(pontoMouse)[1]

    time.sleep(1)
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'w')

    pyautogui.alert("Configuracao feita com sucesso, voce ja pode usar o computador normalmente!")

    return [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]

# ====Menu de configuracoes====
def menuConfiguracoes(listConfiguracoes):
    opc = 1
    while ( opc != 5 and opc > 0 and opc <5):
        print("\nGerenciamento de configuracoes:\n")
        print("1 - Ver configuracoes")
        print("2 - Cadastrar configuracoes")
        print("3 - Atualizar configuracoes")
        print("4 - Deletar configuracoes")
        print("5 - Sair do menu de configuracoes")

        opc = int( input("Digite uma opção: ") )

        if opc == 1:
            mostraTodasConfiguracoes(listConfiguracoes)
            
        elif opc == 2:
            if len(listConfiguracoes) > 0:
                print('Configuracoes ja feitas, selecione a opcao atualizar!')
            else:

                configs = retornarConfigsParaPrograma()

                listConfiguracoes = insereConfiguracoes(listConfiguracoes, configs[0], configs[1], configs[2], configs[3], configs[4], configs[5], configs[6], configs[7], configs[8], configs[9])

                print("Dados inseridos com sucesso!")

                pyautogui.alert("Configuracao feita com sucesso, voce ja pode usar o computador normalmente!")
                pausa()
            
        elif opc == 3:
            configs = retornarConfigsParaPrograma()
            
            listConfiguracoes = alteraConfiguracoes(listConfiguracoes, configs[0], configs[1], configs[2], configs[3], configs[4], configs[5], configs[6], configs[7])
            print("Dados alterados com sucesso!")
            pausa()
            
        elif opc == 4:
            listConfiguracoes = removeConfiguracoes(listConfiguracoes)
            pausa()

        elif opc == 5:
            gravaConfiguracoes(listConfiguracoes)

# ======================================================================

# ======Pega dados do arquivo - INTERFACE====
def recuperaConfiguracoesInterface():
    if (existe_arquivo("configuracoes.txt")):
        arq = open("configuracoes.txt", "r")
        listaConfiguracoes = []
        for linha in arq:
            linha = linha[:len(linha)-1]
            lista = linha.split(";")
            positionArqX = lista[0]
            positionArqY = lista[1]
            positionArq2X = lista[8]
            positionArq2Y = lista[9]
            positionTresPontosX = lista[2]
            positionTresPontosY = lista[3]
            positionFazerDownloadX = lista[4]
            positionFazerDownloadY = lista[5]
            positionEmailX = lista[6]
            positionEmailY = lista[7]
            listaConfiguracoes = [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]
        return listaConfiguracoes
    else:
        return []

# ====Remove configuracoes - INTERFACE====
def removeConfiguracoesInterface(listaConfiguracoes):
    if len(listaConfiguracoes) > 0:
        listaConfiguracoes = []
        return listaConfiguracoes
    else:
        print("Configuracoes não cadastradas!")
        return []

# ====Insere configuracoes - INTERFACE====
def insereConfiguracoesInterface(listaConfiguracoes, positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y):
    if len(listaConfiguracoes) > 0:
        print("Configuracao já cadastrada!")
    else:
        listaConfiguracoes = [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]
        return listaConfiguracoes

# ====Altera configuracoes - INTERFACE====
def alteraConfiguracoesInterface(listaConfiguracoes, positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y):
    if len(listaConfiguracoes) > 0:
        confirma = messagebox.askokcancel("Certeza", "Tem certeza que deseja atualiza-la?: ")
        if confirma == True:
            listaConfiguracoes = [positionArqX, positionArqY, positionTresPontosX, positionTresPontosY, positionFazerDownloadX, positionFazerDownloadY, positionEmailX, positionEmailY, positionArq2X, positionArq2Y]
            return listaConfiguracoes
        else:
            print("Alteração cancelada!")
            return listaConfiguracoes
    else:
        print("Configuracoes não cadastradas!")