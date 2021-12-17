import pyautogui
import time
import pyperclip
import pandas as pd
import os
from gerais import *
from IPython.display import display

BDusuarios = {}
recuperaUsuarios(BDusuarios)

opc = 1

while ( opc != 0 and opc > 0 and opc <3 ):
    menu()

    opc = int( input("Digite uma opção: ") )
        
    if opc == 1:
        menuUsuarios(BDusuarios)
        print("\n\n\n\n")
            
    elif opc == 2:
        print("Codigo da opc 2")
        print("\n\n\n\n")

print("\n\n*** FIM DO PROGRAMA ***\n\n")

os.system('pause')


