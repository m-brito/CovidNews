from gerais import *
from tkinter import messagebox

# ===================Usuarios==============

# ====Verifica se usuario existe====
def existeUsuario(dic,chave):
    if chave in dic.keys():
        return True
    else:
        return False

# ====Insere usuario====
def insereUsuario(dic):
    email = input("Digite o email:")
    municipio = input("Digite o municipio:")
    if existeUsuario(dic, email):
        print("Usuario já cadastrado!")
        pausa()
    else:
        nome = input("Digite o nome: ")
        dic[email]=(nome, municipio)
        print("Dados inseridos com sucesso!")
        pausa()

# ====Exibe um usuario====
def mostraUsuario(dic,chave):
    if existeUsuario(dic,chave):
        dados = dic[chave]
        print(f"Nome: {dados[0]}")
        print(f"Email: {chave}")
        print(f"Município: {dados[1]}")
    else:
        print("Usuario não cadastrada!")

# ====Altera usuario====
def alteraUsuario(dic,chave):
    if existeUsuario(dic,chave):
        mostraUsuario(dic,chave)
        confirma = input("Tem certeza que deseja alterá-lo? (S/N): ").upper()
        if confirma == 'S':
            nome = input("Digite o nome: ")
            municipio = input("Digite o município: ")
            dic[chave]=(nome, municipio)
            print("Dados alterados com sucesso!")
            pausa()
        else:
            print("Alteração cancelada!")
            pausa()
    else:
        print("Usuario não cadastrado!")
        pausa()

# ====Remove um usuario====
def removeUsuario(dic,chave):
    if existeUsuario(dic,chave):
        mostraUsuario(dic,chave)
        confirma = input("Tem certeza que deseja apagar? (S/N): ").upper()
        if confirma == 'S':
            del dic[chave]
            print("Dados apagados com sucesso!")
            pausa()
        else:
            print("Exclusão cancelada!")
            pausa()
    else:
        print("Pessoa não cadastrada!")
        pausa()

# ====Mostra todos usuarios====
def mostraTodosUsuarios(dic):
    print("Relatório: Todas os usuarios\n")
    print("EMAIL - NOME - MUNICÍPIO\n")
    for email in dic:
        tupla = dic[email]
        linha = email + " - " + tupla[0] + " - " + tupla[1]
        print(linha)
    print("")
    pausa()

# ======Grava dados no arquivo=====
def gravaUsuarios(dic):
    arq = open("usuarios.txt", "w")
    for email in dic:
        tupla = dic[email]
        linha = email+";"+tupla[0]+";"+tupla[1]+"\n"
        arq.write(linha)
    arq.close()

# ======Pega dados do arquivo====
def recuperaUsuarios(dic):
    if (existe_arquivo("usuarios.txt")):
        arq = open("usuarios.txt", "r")
        for linha in arq:
            linha = linha[:len(linha)-1]
            lista = linha.split(";")
            nome = lista[1]
            email = lista[0]
            municipio = lista[2]
            dic[email] = (nome, municipio)
            # encode("windows-1252").decode("utf-8")

# def verificaMunicipio(municipio):


# ====Menu de usuarios====
def menuUsuarios(dicUsuarios):
    opc = 0
    while ( opc != 6 ):
        print("\nGerenciamento de usuarios:\n")
        print("1 - Insere Usuario")
        print("2 - Altera Usuario")
        print("3 - Remove Usuario")
        print("4 - Mostra um Usuario")
        print("5 - Mostra todos os Usuarios")
        print("6 - Sair do menu de Usuarios")

        opc = int( input("Digite uma opção: ") )

        if opc == 1:
            insereUsuario(dicUsuarios)
            
        elif opc == 2:
            email = input("Email a ser alterado: ")
            alteraUsuario(dicUsuarios, email)
            
        elif opc == 3:
            email=input("Email a ser removido: ")
            removeUsuario(dicUsuarios, email)
            
        elif opc == 4:
            email=input("Email a ser consultado: ")
            mostraUsuario(dicUsuarios, email)
            pausa()
            
        elif opc == 5:
            mostraTodosUsuarios(dicUsuarios)
            
        elif opc == 6:
            gravaUsuarios(dicUsuarios)

# ====Insere usuario INTERFACE====
def insereUsuarioInterface(dic, email, nome, municipio):
    if existeUsuario(dic, email):
        print("Usuario já cadastrado!")
        messagebox.showinfo("Info", "Usuario já cadastrado!")
    else:
        dic[email]=(nome, municipio)
        print("Dados inseridos com sucesso!")
        messagebox.showinfo("Info", "Dados inseridos com sucesso!")

# ====Remove um usuario INTERFACE====
def removeUsuarioInterface(dic,chave):
    if existeUsuario(dic,chave):
        del dic[chave]
        print("Dados apagados com sucesso!")
        messagebox.showinfo("Info", "Dados apagados com sucesso!")
    else:
        print("Pessoa não cadastrada!")
        messagebox.showinfo("Info", "Pessoa não cadastrada!")

# ====Altera usuario INTERFACE====
def alteraUsuarioInterface(dic, chave, nome, municipio):
    if existeUsuario(dic,chave):
        mostraUsuario(dic,chave)
        dic[chave]=(nome, municipio)
        print("Dados alterados com sucesso!")
        messagebox.showinfo("Info", "Dados alterados com sucesso!")
    else:
        print("Usuario não cadastrado!")
        messagebox.showinfo("Info", "Usuario não cadastrado!")

# ====Exibe um usuario INTERFACE====
def mostraUsuarioInterface(dic,chave):
    if existeUsuario(dic,chave):
        dados = dic[chave]
        return(dados[0], chave, dados[1])
    else:
        return(False, False, False)