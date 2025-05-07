import time
import os
import win32com.client

import tkinter as tk
from tkinter import ttk
from tkinter import *

import PyPolitica
import PyVerifica





# ----------------------------- Primeira parte do codigo --------------------

PyPolitica.executar_powershell_alteracao_politicas

time.sleep(6)

# executa o PyVerifica.py que é o arquivo que verifica se o choco esta instalado de fato no sistema
PyVerifica.Verificar_Choco

# ----------------------------- Primeira parte do codigo --------------------




# Caminho do script PowerShell
comando_choco = r"C:\Users\vboxuser\Documents\github\ProjetoCHOCO\ChocoInstall.ps1"

# Cria o objeto Shell.Application para executar com privilégios
shell = win32com.client.Dispatch("Shell.Application")

# Função para executar um script PowerShell elevado
def executar_powershell():
    shell.ShellExecute(
        "powershell.exe",
        f"-ExecutionPolicy Bypass -File \"{comando_choco}\"",
        "",
        "runas",
        1
    )

# Cria a janela principal
janela = Tk()
janela.title("Programa com Um Click")

# Labels de orientação
Label(janela, text="Escolha abaixo os programas que deseja instalar").grid(column=0, row=0, padx=10, pady=5)
Label(janela, text="---------------------------------------").grid(column=0, row=1, padx=10, pady=5)


# Variável de controle do Checkbutton
check_varGoogle = BooleanVar()
check_varpdf24 = BooleanVar()
check_varWinrar = BooleanVar()
check_varJava = BooleanVar()
check_varAnydesk = BooleanVar()
check_varlibreoffice = BooleanVar()

# Label para mostrar o estado
texto_programas = Label(janela, text="Programas Escolhidos: \n")
texto_programas.grid(column=2, row=1, padx=10, pady=5)

# texto com os programas a serem instalado
txt_program_google = Label(janela, text="Google Chrome Desmarcado [X]")
txt_program_google.grid(column=2, row=2, padx=2, pady=1)

txt_program_pdf24 = Label(janela, text="PDF-24 Desmarcado [X]")
txt_program_pdf24.grid(column=2, row=3, padx=2, pady=1)

txt_program_winrar = Label(janela, text="Winrar Desmarcado [X]")
txt_program_winrar.grid(column=2, row=4, padx=2, pady=1)

txt_program_javaruntime = Label(janela, text="Java Desmarcado [X]")
txt_program_javaruntime.grid(column=2, row=5, padx=2, pady=1)

txt_program_anydesk = Label(janela, text="AnyDesk Desmarcado [X]")
txt_program_anydesk.grid(column=2, row=6, padx=2, pady=1)

txt_program_libreoffice = Label(janela, text="libreoffice Desmarcado [X]")
txt_program_libreoffice.grid(column=2, row=7, padx=2, pady=1)

# escreve na tela quais programas estao marcado para ser instalado
def mostrar_estado():
    estado1 = "Google Chrome Marcado" if check_varGoogle.get() else "Google Chrome Desmarcado [X]\n"
    txt_program_google.config(text=f"{estado1}")
    
    estado1 = "PDF-24 Marcado" if check_varpdf24.get() else "PDF-24 Desmarcado [X]\n"
    txt_program_pdf24.config(text=f"{estado1}")

    estado1 = "Winrar Marcado" if check_varWinrar.get() else "Winrar Desmarcado [X] \n"
    txt_program_winrar.config(text=f"{estado1}")

    estado1 = "Java Marcado" if check_varJava.get() else "Java Desmarcado [X] \n"
    txt_program_javaruntime.config(text=f"{estado1}")

    estado1 = "AnyDesk Marcado" if check_varAnydesk.get() else "AnyDesk Desmarcado [X] \n"
    txt_program_anydesk.config(text=f"{estado1}")

    estado1 = "Libre Office Marcado" if check_varlibreoffice.get() else "Libre Office Desmarcado [X] \n"
    txt_program_libreoffice.config(text=f"{estado1}")

    # Verifica se os programas estao marcado como TRUE, se sim escreve no arquivo ChocoInstall.ps1, o comando da instalaçao
    # instalador do google
    if check_varGoogle.get() == True:
        with open("ChocoInstall.ps1", "w", encoding="utf-8") as arquivo:
            arquivo.write("choco install googlechrome -y --force\n")
    else:
        with open("ChocoInstall.ps1", "w", encoding="utf-8") as arquivo:
            arquivo.write("\n")

    # instalador do pdf-24
    if check_varpdf24.get() == True:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("choco install pdf24 -y --force\n")
    else:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("\n")
    
    # instalador do winrar
    if check_varWinrar.get() == True:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("choco install winrar -y --force\n")
    else:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("\n")

    # instalador do java
    if check_varJava.get() == True:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("choco install javaruntime -y --force\n")
    else: 
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("\n")

    # instalador do AnyDesk
    if check_varAnydesk.get() == True:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("choco install anydesk -y --force\n")
    else: 
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("\n")

     # instalador do AnyDesk
    if check_varlibreoffice.get() == True:
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("choco install libreoffice-fresh -y --force\n")
    else: 
        with open("ChocoInstall.ps1", "a", encoding="utf-8") as arquivo:
            arquivo.write("\n")

    

# check 01
# cria o Checkbutton vinna tela
checkbox1 = Checkbutton(janela, text="GOOGLE", variable=check_varGoogle, command=mostrar_estado)
checkbox1.grid(column=0, row=2, padx=2, pady=2)

# check 02
# Checkbutton vinculado à variável
checkbox2 = Checkbutton(janela, text="PDF 24", variable=check_varpdf24, command=mostrar_estado)
checkbox2.grid(column=0, row=3, padx=2, pady=2)

# check 03
# Checkbutton vinculado à variável
checkbox3 = Checkbutton(janela, text="Winrar", variable=check_varWinrar, command=mostrar_estado)
checkbox3.grid(column=0, row=4, padx=2, pady=2)

# check 04
# Checkbutton vinculado à variável
checkbox4 = Checkbutton(janela, text="Java", variable=check_varJava, command=mostrar_estado)
checkbox4.grid(column=0, row=5, padx=2, pady=2)

# check 05
# Checkbutton vinculado à variável
checkbox4 = Checkbutton(janela, text="AnyDesk", variable=check_varAnydesk, command=mostrar_estado)
checkbox4.grid(column=0, row=6, padx=2, pady=2)

# check 06
# Checkbutton vinculado à variável
checkbox4 = Checkbutton(janela, text="Libre Office", variable=check_varlibreoffice, command=mostrar_estado)
checkbox4.grid(column=0, row=7, padx=2, pady=2)



# ---------------------------------- ações --------------------------------
# Botão de instalação, só habilita se o checkbox estiver marcado
def on_install_click():
        executar_powershell()



install_btn = Button(janela, text="Clique para instalar", command=on_install_click)
install_btn.grid(pady=20)



# Inicia o loop da interface
janela.mainloop()
