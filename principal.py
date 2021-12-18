import time
import os

from datetime import datetime

from gerais import *
from usuario import *
from configuracoes import *

BDusuarios = {}
BDrelatorio = {}
BDconfiguracoes = []
recuperaUsuarios(BDusuarios)
BDconfiguracoes = recuperaConfiguracoes()

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
            print("\n\n\n\n")

    elif opc == 3:
        BDconfiguracoes = recuperaConfiguracoes()
        menuConfiguracoes(BDconfiguracoes)
        print("\n\n\n\n")

print("\n\n*** FIM DO PROGRAMA ***\n\n")

os.system('pause')