from datetime import datetime as date

def menu():
    print('1 - Gerenciar Usuarios')
    print('2 - Gerenciar Relatorios')
    print('3 - Gerenciar Confirguracoes do app')

    print('0 - Sair')

def existe_arquivo(nome):
    import os
    if os.path.exists(nome):
        return True
    else:
        return False

def dataAtualF():
    dataAtual = date.today()
    dataFormatada = dataAtual.strftime('%d/%m/%Y')
    return dataFormatada
def dataAtualFArquivo():
    dataAtual = date.today()
    dataFormatada = dataAtual.strftime('%d-%m-%Y')
    return dataFormatada
def mesAtual():
    dataAtual = date.today()
    dataFormatada = dataAtual.strftime('%m')
    return dataFormatada
def anoAtual():
    dataAtual = date.today()
    dataFormatada = dataAtual.strftime('%Y')
    return dataFormatada

def pausa():
    input("Tecle <ENTER> para continuar...\n")