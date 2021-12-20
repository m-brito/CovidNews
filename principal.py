from ctypes import BigEndianStructure
import time
import os

from datetime import datetime
from tkinter import *
from tkinter import ttk

from pyautogui import scroll

from gerais import *
from usuario import *
from relatorio import *
from configuracoes import *

BDusuarios = {}
BDrelatorio = {}
BDconfiguracoes = []
recuperaUsuarios(BDusuarios)
recuperaRelatorios(BDrelatorio)
BDconfiguracoes = recuperaConfiguracoes()

# ========================TESTES INTERFACE===============================
class FuncsUsuarios():
    def limparTela(self):
        self.emailEntry.delete(0, END)
        self.nomeEntry.delete(0, END)

    def variaveis(self):
        self.campoEmail = self.emailEntry.get()
        self.campoNome = self.nomeEntry.get()

    def sair(self):
        self.janelaUsuarios.destroy()

    def addUsuario(self):
        self.variaveis()
        insereUsuarioInterface(BDusuarios, self.campoEmail, self.campoNome)
        self.mostrarUsuarios()
        self.limparTela()
        gravaUsuarios(BDusuarios)

    def mostrarUsuarios(self):
        self.listaUsuarios.delete(*self.listaUsuarios.get_children())
        for email in BDusuarios:
            tupla = BDusuarios[email]
            self.listaUsuarios.insert("", END, values=(email, tupla))

    def duploClique(self, event):
        self.limparTela()
        self.listaUsuarios.selection()
        for x in self.listaUsuarios.selection():
            col1, col2 = self.listaUsuarios.item(x, 'values')
            self.emailEntry.insert(END, col1)
            self.nomeEntry.insert(END, col2)

    def deletarUsuario(self):
        self.variaveis()
        removeUsuarioInterface(BDusuarios, self.campoEmail)
        self.mostrarUsuarios()
        self.limparTela()
        gravaUsuarios(BDusuarios)

    def alterarUsuario(self):
        self.variaveis()
        alteraUsuarioInterface(BDusuarios, self.campoEmail, self.campoNome)
        self.mostrarUsuarios()
        self.limparTela()
        gravaUsuarios(BDusuarios)

    def buscarUsuario(self):
        self.variaveis()
        respNome, respEmail = mostraUsuarioInterface(BDusuarios, self.campoEmail)
        self.limparTela()
        self.mostrarUsuarios()
        if respNome != False and respEmail != False:
            self.emailEntry.insert(END, respEmail)
            self.nomeEntry.insert(END, respNome)

class AplicationRelatorio():
    def __init__(self):
        self.janelaRelatorios = Tk()
        self.telaRelatorios()
        self.menu()
        self.janelaRelatorios.mainloop()
    
    def telaUsuarios(self):
        self.janelaRelatorios.destroy()
        AplicationUsuarios()
        print('telinha de usuarios')

    def telaRelatorios(self):
        self.janelaRelatorios.title('Gerenciamento de relatorios')
        self.janelaRelatorios.configure(background='#D8F3DC')
        self.janelaRelatorios.geometry('700x500')
        self.janelaRelatorios.resizable(True, True)
        self.janelaRelatorios.maxsize(width=900, height=700)
        self.janelaRelatorios.minsize(width=500, height=500)
        print('telinha de relatorios')

    def telaConfiguracoes(self):
        self.janelaRelatorios.destroy()
        AplicationConfiguracoes()
        print('telinha de configuracoes')

    def menu(self):
        menuBar = Menu(self.janelaRelatorios)
        self.janelaRelatorios.config(menu=menuBar)
        fileMenu = Menu(menuBar)
        fileMenu2 = Menu(menuBar)
        
        menuBar.add_cascade(label= 'Opcoes', menu=fileMenu)
        menuBar.add_cascade(label= 'Sobre', menu=fileMenu2)

        fileMenu.add_command(label='Usuarios', command= self.telaUsuarios)
        fileMenu.add_command(label='Relatorios', command= self.telaRelatorios)
        fileMenu.add_command(label='Configuracoes', command= self.telaConfiguracoes)

class AplicationConfiguracoes():
    def __init__(self):
        self.janelaConfiguracoes = Tk()
        self.telaConfiguracoes()
        self.menu()
        self.janelaConfiguracoes.mainloop()
    
    def telaUsuarios(self):
        self.janelaConfiguracoes.destroy()
        AplicationUsuarios()
        print(1)

    def telaRelatorios(self):
        self.janelaConfiguracoes.destroy()
        AplicationRelatorio()
        print('telinha de relatorios')

    def telaConfiguracoes(self):
        self.janelaConfiguracoes.title('Gerenciamento de configuracoes')
        self.janelaConfiguracoes.configure(background='#D8F3DC')
        self.janelaConfiguracoes.geometry('700x500')
        self.janelaConfiguracoes.resizable(True, True)
        self.janelaConfiguracoes.maxsize(width=900, height=700)
        self.janelaConfiguracoes.minsize(width=500, height=500)
        print('telinha de configuracoes')

    def menu(self):
        menuBar = Menu(self.janelaConfiguracoes)
        self.janelaConfiguracoes.config(menu=menuBar)
        fileMenu = Menu(menuBar)
        fileMenu2 = Menu(menuBar)
        
        menuBar.add_cascade(label= 'Opcoes', menu=fileMenu)
        menuBar.add_cascade(label= 'Sobre', menu=fileMenu2)

        fileMenu.add_command(label='Usuarios', command= self.telaUsuarios)
        fileMenu.add_command(label='Relatorios', command= self.telaRelatorios)
        fileMenu.add_command(label='Configuracoes', command= self.telaConfiguracoes)
class AplicationUsuarios(FuncsUsuarios):
    def __init__(self):
        self.janelaUsuarios = Tk()
        self.telaUsuarios()
        self.framesDaTela()
        self.widgetsFrame1()
        self.listaFrame2()
        self.mostrarUsuarios()
        self.menu()
        self.janelaUsuarios.mainloop()

    def telaRelatorios(self):
        self.janelaUsuarios.destroy()
        AplicationRelatorio()
        print('telinha relatorios')

    def telaConfiguracoes(self):
        self.janelaUsuarios.destroy()
        AplicationConfiguracoes()
        print('telinha configuracoes')

    def telaUsuarios(self):
        self.janelaUsuarios.title('Gerenciamento de usuarios')
        self.janelaUsuarios.configure(background='#D8F3DC')
        self.janelaUsuarios.geometry('700x500')
        self.janelaUsuarios.resizable(True, True)
        self.janelaUsuarios.maxsize(width=900, height=700)
        self.janelaUsuarios.minsize(width=500, height=500)
        print('telinha de usuarios')

    def framesDaTela(self):
        self.frame1 = Frame(self.janelaUsuarios, bd=4, bg='white', highlightbackground='#95D5B2', highlightthickness=2)
        self.frame1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.36)

        self.frame2 = Frame(self.janelaUsuarios, bd=4, bg='white', highlightbackground='#95D5B2', highlightthickness=2)
        self.frame2.place(relx=0.02, rely=0.4, relwidth=0.96, relheight=0.56)
    
    def widgetsFrame1(self):
        self.btnLimpar = Button(self.frame1, text='Limpar', bg='#EBEBEB', command=self.limparTela)
        self.btnLimpar.place(relx=0.4, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnBuscar = Button(self.frame1, text='Buscar', bg='#EBEBEB', command=self.buscarUsuario)
        self.btnBuscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnNovo = Button(self.frame1, text='Novo', bg='#EBEBEB', command=self.addUsuario)
        self.btnNovo.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnAlterar = Button(self.frame1, text='Alterar', bg='#EBEBEB', command=self.alterarUsuario)
        self.btnAlterar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnApagar = Button(self.frame1, text='Apagar', bg='#EBEBEB', command=self.deletarUsuario)
        self.btnApagar.place(relx=0.9, rely=0.1, relwidth=0.1, relheight=0.15)

        # =========================================================

        self.lblEmail = Label(self.frame1, text='Email', bg='white')
        self.lblEmail.place(relx=0.01, rely=0.01)
        self.emailEntry = Entry(self.frame1, bg='#EBEBEB')
        self.emailEntry.place(relx=0.01, rely=0.11, relheight=0.15, relwidth=0.3)

        self.lblNome = Label(self.frame1, text='Nome', bg='white')
        self.lblNome.place(relx=0.01, rely=0.4)
        self.nomeEntry = Entry(self.frame1, bg='#EBEBEB')
        self.nomeEntry.place(relx=0.01, rely=0.5, relheight=0.15, relwidth=0.989)

    def listaFrame2(self):
        self.listaUsuarios = ttk.Treeview(self.frame2, height=4, columns=('col1', 'col2'))
        self.listaUsuarios.heading('#0', text='')
        self.listaUsuarios.heading('#1', text='Email')
        self.listaUsuarios.heading('#2', text='Nome')

        self.listaUsuarios.column('#0', width=1)
        self.listaUsuarios.column('#1', width=350)
        self.listaUsuarios.column('#2', width=250)

        self.listaUsuarios.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.8)

        self.scrollLista = Scrollbar(self.frame2, orient='vertical')
        self.listaUsuarios.configure(yscroll=self.scrollLista.set)
        self.scrollLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.8)
        self.listaUsuarios.bind('<Double-1>', self.duploClique)

    def menu(self):
        menuBar = Menu(self.janelaUsuarios)
        self.janelaUsuarios.config(menu=menuBar)
        fileMenu = Menu(menuBar)
        fileMenu2 = Menu(menuBar)
        
        menuBar.add_cascade(label= 'Opcoes', menu=fileMenu)
        menuBar.add_cascade(label= 'Sobre', menu=fileMenu2)

        fileMenu.add_command(label='Usuarios', command= self.telaUsuarios)
        fileMenu.add_command(label='Relatorios', command= self.telaRelatorios)
        fileMenu.add_command(label='Configuracoes', command= self.telaConfiguracoes)
        fileMenu.add_command(label='Sair', command= self.sair)
        fileMenu2.add_command(label='Limpa Usuarios', command= self.limparTela)

AplicationUsuarios()

# =======================================================================


def verificarConfiguracoes(msgErro):
    if len(BDconfiguracoes) == 0:
        print(f'\n====================================================================\n{msgErro}\n====================================================================\n')
        return False
    else:
        return True

opc = 1

while ( opc != 0 and opc > 0 and opc <4 ):
    BDconfiguracoes = recuperaConfiguracoes()
    menu()

    opc = int( input("Digite uma opção: ") )
        
    if opc == 1:
        menuUsuarios(BDusuarios)
        print("\n\n\n\n")
            
    elif opc == 2:
        BDconfiguracoes = recuperaConfiguracoes()
        if verificarConfiguracoes('Voce não consegue acessar aqui sem antes configurar o app!!!') == True:
            menuRelatorios(BDrelatorio)
            print("\n\n\n\n")

    elif opc == 3:
        BDconfiguracoes = recuperaConfiguracoes()
        menuConfiguracoes(BDconfiguracoes)
        print("\n\n\n\n")

print("\n\n*** FIM DO PROGRAMA ***\n\n")

os.system('pause')