from ctypes import BigEndianStructure
from io import DEFAULT_BUFFER_SIZE
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

import webbrowser
from pathlib import Path

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
        self.municipioEntry.delete(0, END)

    def variaveis(self):
        self.campoEmail = self.emailEntry.get()
        self.campoNome = self.nomeEntry.get()
        self.campoMunicipio = self.municipioEntry.get()

    def sair(self):
        self.janelaUsuarios.destroy()

    def addUsuario(self):
        self.variaveis()
        email = self.campoEmail
        nome = self.campoNome
        municipio = self.campoMunicipio
        if len(str(email).replace(" ", "")) == 0 or len(str(nome).replace(" ", "")) == 0 or len(str(municipio).replace(" ", "")) == 0:
            messagebox.showinfo("Info", "Preencha todos os campos!")
        else:
            insereUsuarioInterface(BDusuarios, self.campoEmail, self.campoNome, self.campoMunicipio)
            self.mostrarUsuarios()
            self.limparTela()
            gravaUsuarios(BDusuarios)

    def mostrarUsuarios(self):
        self.listaUsuarios.delete(*self.listaUsuarios.get_children())
        for email in BDusuarios:
            tupla = BDusuarios[email]
            self.listaUsuarios.insert("", END, values=(email, tupla[0], tupla[1]))

    def duploClique(self, event):
        self.limparTela()
        self.listaUsuarios.selection()
        for x in self.listaUsuarios.selection():
            col1, col2, col3 = self.listaUsuarios.item(x, 'values')
            self.emailEntry.insert(END, col1)
            self.nomeEntry.insert(END, col2)
            self.municipioEntry.insert(END, col3)

    def deletarUsuario(self):
        self.variaveis()
        removeUsuarioInterface(BDusuarios, self.campoEmail)
        self.mostrarUsuarios()
        self.limparTela()
        gravaUsuarios(BDusuarios)

    def alterarUsuario(self):
        self.variaveis()
        email = self.campoEmail
        nome = self.campoNome
        municipio = self.campoMunicipio
        if len(str(email).replace(" ", "")) == 0 or len(str(nome).replace(" ", "")) == 0 or len(str(municipio).replace(" ", "")) == 0:
            messagebox.showinfo("Info", "Preencha todos os campos!")
        else:
            alteraUsuarioInterface(BDusuarios, self.campoEmail, self.campoNome, self.campoMunicipio)
            self.mostrarUsuarios()
            self.limparTela()
            gravaUsuarios(BDusuarios)

    def buscarUsuario(self):
        self.variaveis()
        respNome, respEmail, respMunicipio = mostraUsuarioInterface(BDusuarios, self.campoEmail)
        self.limparTela()
        self.mostrarUsuarios()
        if respNome != False and respEmail != False and respMunicipio != False:
            self.emailEntry.insert(END, respEmail)
            self.nomeEntry.insert(END, respNome)
            self.municipioEntry.insert(END, respMunicipio)

    def abrirMunicipios(self):
        self.variaveis()
        path = os.getcwd()
        diretorio = f"{path}\\Municipios\\Municipios de SP.pdf"
        if Path(diretorio).is_file():
            webbrowser.open(diretorio)
        else:
            print("O arquivo não existe!!!")

class FuncConfiguracoes():
    def mostrarConfigs(self):
        BDconfiguracoes = recuperaConfiguracoesInterface()
        if(len(BDconfiguracoes)>=1):
            self.editarLblPA(BDconfiguracoes[0], BDconfiguracoes[1])
            self.editarLblPA2(BDconfiguracoes[8], BDconfiguracoes[9])
            self.editarLblPTP(BDconfiguracoes[2], BDconfiguracoes[3])
            self.editarLblPFD(BDconfiguracoes[4], BDconfiguracoes[5])
            self.editarLblPE(BDconfiguracoes[6], BDconfiguracoes[7])
            self.editarLblDiretorioDonwloads(BDconfiguracoes[10])
        else:
            self.resetarLbl()

    def editarLblPA(self, px, py):
        self.lblArquivo['text'] = f'Posição do arquivo 1: ({px}, {py})'

    def editarLblPA2(self, px, py):
        self.lblArquivo2['text'] = f'Posição do arquivo 2: ({px}, {py})'

    def editarLblPTP(self, px, py):
        self.lblTresPontos['text'] = f'Posição dos três pontos: ({px}, {py})'

    def editarLblPFD(self, px, py):
        self.lblFzrDownload['text'] = f'Posição do "Fazer Donwload": ({px}, {py})'

    def editarLblPE(self, px, py):
        self.lblEmail['text'] = f'Posição do "Escrever": ({px}, {py})'

    def editarLblDiretorioDonwloads(self, diretorio):
        self.lblDiretorioDonwloads['text'] = f'Diretorio de Donwloads: {diretorio}'

    def resetarLbl(self):
        self.lblArquivo['text'] = 'Posição do arquivo 1: (Nenhum, Nenhum)'
        self.lblArquivo2['text'] = 'Posição do arquivo 2: (Nenhum, Nenhum)'
        self.lblTresPontos['text'] = 'Posição dos três pontos: (Nenhum, Nenhum)'
        self.lblFzrDownload['text'] = 'Posição do "Fazer Donwload": (Nenhum, Nenhum)'
        self.lblEmail['text'] = 'Posição do "Escrever": (Nenhum, Nenhum)'
        self.lblDiretorioDonwloads['text'] = 'Diretorio de Donwloads: -------'

    def apagarConfigs(self):
        self.BDconfiguracoes = removeConfiguracoesInterface(self.BDconfiguracoes)
        gravaConfiguracoes(self.BDconfiguracoes)
        self.mostrarConfigs()

    def cadConfiguracoes(self):
        if len(self.BDconfiguracoes) <= 0:
            configs = retornarConfigsParaPrograma()

            self.BDconfiguracoes = insereConfiguracoes(self.BDconfiguracoes, configs[0], configs[1], configs[2], configs[3], configs[4], configs[5], configs[6], configs[7], configs[8], configs[9], configs[10])
            gravaConfiguracoes(self.BDconfiguracoes)
            self.BDconfiguracoes = recuperaConfiguracoes()
            self.mostrarConfigs()

            print("Dados inseridos com sucesso!")
        else:
            print('Configuracoes ja feitas! Tente atualizar!')
    
    def atuaConfiguracoes(self):
        if len(self.BDconfiguracoes) >=1:
            configs = retornarConfigsParaPrograma()
                
            self.BDconfiguracoes = alteraConfiguracoesInterface(self.BDconfiguracoes, configs[0], configs[1], configs[2], configs[3], configs[4], configs[5], configs[6], configs[7])
            gravaConfiguracoes(self.BDconfiguracoes)
            self.mostrarConfigs()
        else:
            print("Nenhuma configuracao salva para atualizar!!")
        
class FuncRelatorio():
    def limparTela(self):
        self.dataEntry.delete(0, END)
        self.lblEnviou['text'] = f'Enviou email: ---'
        self.lblAtualizou['text'] = f'Atualizou: ---'

    def variaveis(self):
        self.campoData = self.dataEntry.get()

    def duploClique(self, event):
        self.limparTela()
        self.listaRelatorio.selection()
        for x in self.listaRelatorio.selection():
            col1, col2, col3 = self.listaRelatorio.item(x, 'values')
            self.dataEntry.insert(END, col1)
        self.lblEnviou['text'] = f'Enviou email: {col2}'
        self.lblAtualizou['text'] = f'Atualizou: {col3}'

    def mostrarRelatorios(self):
        self.listaRelatorio.delete(*self.listaRelatorio.get_children())
        for data in BDrelatorio:
            tupla = BDrelatorio[data]
            self.listaRelatorio.insert("", END, values=(data, tupla[0], tupla[1]))

    def criandoRelatorio(self):
        try:
            data, enviou, editou = criarRelatorioInterface()
            insereRelatorioInterface(BDrelatorio, data, enviou, editou)
            gravaRelatorios(BDrelatorio)
            recuperaRelatorios(BDrelatorio)
            self.mostrarRelatorios()
        except:
            print("Tivemos alguns problemas para criar o relatorio")
            messagebox.showinfo("Info", "Tivemos alguns problemas para criar o relatorio!")

    def criaRelatorio(self):
        dataArq = str(dataAtualFArquivo())
        path = os.getcwd()
        diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{dataArq}.pdf"
        if Path(diretorio).is_file():
            okcancel = messagebox.askokcancel("Atenção", "Relatorio ja cadastrado, se proseguir ira atualizar um relatorio ja existente")
            if okcancel == True:
                self.criandoRelatorio()
        else:
            self.criandoRelatorio()

        print("Dados inseridos com sucesso!")
    
    def deletarRelatorio(self):
        self.variaveis()
        removeRelatorioInterface(BDrelatorio, self.campoData)
        self.mostrarRelatorios()
        self.limparTela()
        gravaRelatorios(BDrelatorio)
    
    def buscarRelatorio(self):
        self.variaveis()
        respData, respEnviou, respAtualizou = mostraRelatorioInterface(BDrelatorio, self.campoData)
        self.limparTela()
        self.mostrarRelatorios()
        if respData != False and respEnviou != False and respAtualizou != False:
            self.dataEntry.insert(END, respData)
            self.lblEnviou['text'] = f'Enviou email: {respEnviou}'
            self.lblAtualizou['text'] = f'Atualizou: {respAtualizou}'

    def abrirRelatorio(self):
        self.variaveis()
        data = str(self.campoData).replace('/', '-')
        path = os.getcwd()
        diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{data}.pdf"
        if Path(diretorio).is_file():
            webbrowser.open(diretorio)
        else:
            print("O arquivo não existe!!!")

    def enviarRelatorio(self):
        self.variaveis()
        data = str(self.campoData)
        dataArq = str(self.campoData).replace('/', '-')
        path = os.getcwd()
        diretorio = f"{path}\\RelatoriosEgraficos\\relatorio-{dataArq}.pdf"
        if Path(diretorio).is_file():
            enviarRelatorioEmail(dataArq, data)
            respData, respEnviou, respAtualizou = mostraRelatorioInterface(BDrelatorio, data)
            alteraRelatorioInterface(BDrelatorio, respData, "SIM", respAtualizou)
            gravaRelatorios(BDrelatorio)
            recuperaRelatorios(BDrelatorio)
            self.mostrarRelatorios()
            print("Relatorio enviado com sucesso")
        else:
            print("O arquivo não existe!!!")


class AplicationRelatorio(FuncRelatorio):
    def __init__(self):
        self.janelaRelatorios = Tk()
        self.telaRelatorios()
        self.framesDaTela()
        self.widgetsFrame1()
        self.listaFrame2()
        self.mostrarRelatorios()
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
        self.janelaRelatorios.minsize(width=700, height=500)
        print('telinha de relatorios')
    
    def framesDaTela(self):
        self.frame1 = Frame(self.janelaRelatorios, bd=4, bg='white', highlightbackground='#95D5B2', highlightthickness=2)
        self.frame1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.36)

        self.frame2 = Frame(self.janelaRelatorios, bd=4, bg='white', highlightbackground='#95D5B2', highlightthickness=2)
        self.frame2.place(relx=0.02, rely=0.4, relwidth=0.96, relheight=0.56)

    def widgetsFrame1(self):
        self.btnLimpar = Button(self.frame1, text='Limpar', bg='#EBEBEB', command=self.limparTela)
        self.btnLimpar.place(relx=0.4, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnBuscar = Button(self.frame1, text='Buscar', bg='#EBEBEB', command=self.buscarRelatorio)
        self.btnBuscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnNovo = Button(self.frame1, text='Criar', bg='#EBEBEB', command=self.criaRelatorio)
        self.btnNovo.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnApagar = Button(self.frame1, text='Apagar', bg='#EBEBEB', command=self.deletarRelatorio)
        self.btnApagar.place(relx=0.9, rely=0.1, relwidth=0.1, relheight=0.15)

        self.btnAbrir = Button(self.frame1, text='Abrir', bg='#EBEBEB', command=self.abrirRelatorio)
        self.btnAbrir.place(relx=0.9, rely=0.28, relwidth=0.1, relheight=0.15)

        self.btnEnviar = Button(self.frame1, text='Enviar', bg='#EBEBEB', command=self.enviarRelatorio)
        self.btnEnviar.place(relx=0.8, rely=0.28, relwidth=0.1, relheight=0.15)

        # =========================================================

        self.lblData = Label(self.frame1, text='Data', bg='white')
        self.lblData.place(relx=0.01, rely=0.01)
        self.dataEntry = Entry(self.frame1, bg='#EBEBEB')
        self.dataEntry.place(relx=0.01, rely=0.11, relheight=0.15, relwidth=0.3)

        self.lblEnviou = Label(self.frame1, text='Enviou email: ---', bg='white')
        self.lblEnviou.place(relx=0.01, rely=0.31)

        self.lblAtualizou = Label(self.frame1, text='Atualizou: ---', bg='white')
        self.lblAtualizou.place(relx=0.01, rely=0.51)
    
    def listaFrame2(self):
        self.listaRelatorio = ttk.Treeview(self.frame2, height=4, columns=('col1', 'col2', 'col3'))
        self.listaRelatorio.heading('#0', text='')
        self.listaRelatorio.heading('#1', text='Data')
        self.listaRelatorio.heading('#2', text='Enviou email')
        self.listaRelatorio.heading('#3', text='Atualizou')

        self.listaRelatorio.column('#0', width=1)
        self.listaRelatorio.column('#1', width=200)
        self.listaRelatorio.column('#2', width=200)
        self.listaRelatorio.column('#2', width=200)

        self.listaRelatorio.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.8)

        self.scrollLista = Scrollbar(self.frame2, orient='vertical')
        self.listaRelatorio.configure(yscroll=self.scrollLista.set)
        self.scrollLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.8)
        self.listaRelatorio.bind('<Double-1>', self.duploClique)

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

class AplicationConfiguracoes(FuncConfiguracoes):
    def __init__(self):
        self.BDconfiguracoes = list(BDconfiguracoes)
        self.janelaConfiguracoes = Tk()
        self.telaConfiguracoes()
        self.framesDaTela()
        self.widgetsFrame1()
        self.mostrarConfigs()
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
        self.janelaConfiguracoes.minsize(width=700, height=500)
        print('telinha de configuracoes')
    
    def framesDaTela(self):
        self.frame1 = Frame(self.janelaConfiguracoes, bd=4, bg='white', highlightbackground='#95D5B2', highlightthickness=2)
        self.frame1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)

    def widgetsFrame1(self):
        self.btnNovo = Button(self.frame1, text='Cadastrar', bg='#EBEBEB', command=self.cadConfiguracoes)
        self.btnNovo.place(relx=0.01, rely=0.01, relwidth=0.98, relheight=0.10)

        self.btnAtualizar = Button(self.frame1, text='Atualizar', bg='#EBEBEB', command=self.atuaConfiguracoes)
        self.btnAtualizar.place(relx=0.01, rely=0.13, relwidth=0.98, relheight=0.10)

        self.btnApagar = Button(self.frame1, text='Apagar', bg='#EBEBEB', command=self.apagarConfigs)
        self.btnApagar.place(relx=0.01, rely=0.25, relwidth=0.98, relheight=0.10)

        self.lblArquivo = Label(self.frame1, text='Posicao do arquivo 1: (Nenhum, Nenhum)', bg='white')
        self.lblArquivo.place(relx=0.35, rely=0.45)

        self.lblArquivo2 = Label(self.frame1, text='Posicao do arquivo 2: (Nenhum, Nenhum)', bg='white')
        self.lblArquivo2.place(relx=0.35, rely=0.5)

        self.lblTresPontos = Label(self.frame1, text='Posicao dos tres pontos: (Nenhum, Nenhum)', bg='white')
        self.lblTresPontos.place(relx=0.35, rely=0.55)

        self.lblFzrDownload = Label(self.frame1, text='Posicao do "Fazer download": (Nenhum, Nenhum)', bg='white')
        self.lblFzrDownload.place(relx=0.35, rely=0.6)

        self.lblEmail = Label(self.frame1, text='Posicao do "Escrever": (Nenhum, Nenhum)', bg='white')
        self.lblEmail.place(relx=0.35, rely=0.65)

        self.lblAvisoFixo = Label(self.frame1, text='Outras configuracoes:', bg='white')
        self.lblAvisoFixo.place(relx=0.01, rely=0.86)

        self.lblDiretorioDonwloads = Label(self.frame1, text='Diretorio de Donwloads: ---', bg='white')
        self.lblDiretorioDonwloads.place(relx=0.01, rely=0.90)

        self.lblLinkDrive = Label(self.frame1, text='Link do arquivo usado: https://drive.google.com/drive/folders/1dSPGvocU1Tp-KobogpWfCBJF8QXEcSCY?usp=sharing', bg='white')
        self.lblLinkDrive.place(relx=0.01, rely=0.94)

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
        self.janelaUsuarios.minsize(width=700, height=500)
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

        self.btnVerMunicipios = Button(self.frame1, text='Ver Municipios', bg='#EBEBEB', command=self.abrirMunicipios)
        self.btnVerMunicipios.place(relx=0.85, rely=0.69, relwidth=0.15, relheight=0.15)

        # =========================================================

        self.lblEmail = Label(self.frame1, text='Email', bg='white')
        self.lblEmail.place(relx=0.01, rely=0.01)
        self.emailEntry = Entry(self.frame1, bg='#EBEBEB')
        self.emailEntry.place(relx=0.01, rely=0.11, relheight=0.15, relwidth=0.3)

        self.lblNome = Label(self.frame1, text='Nome', bg='white')
        self.lblNome.place(relx=0.01, rely=0.3)
        self.nomeEntry = Entry(self.frame1, bg='#EBEBEB')
        self.nomeEntry.place(relx=0.01, rely=0.4, relheight=0.15, relwidth=0.989)

        self.lblMunicipio = Label(self.frame1, text='Município', bg='white')
        self.lblMunicipio.place(relx=0.01, rely=0.6)
        self.municipioEntry = Entry(self.frame1, bg='#EBEBEB')
        self.municipioEntry.place(relx=0.01, rely=0.7, relheight=0.15, relwidth=0.829)

    def listaFrame2(self):
        self.listaUsuarios = ttk.Treeview(self.frame2, height=4, columns=('col1', 'col2', 'col3'))
        self.listaUsuarios.heading('#0', text='')
        self.listaUsuarios.heading('#1', text='Email')
        self.listaUsuarios.heading('#2', text='Nome')
        self.listaUsuarios.heading('#3', text='Município')

        self.listaUsuarios.column('#0', width=1)
        self.listaUsuarios.column('#1', width=300)
        self.listaUsuarios.column('#2', width=150)
        self.listaUsuarios.column('#3', width=150)

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