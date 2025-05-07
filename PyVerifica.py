
import time
import os
import win32com.client


file_ChocoVerifica = r"C:\Users\vboxuser\Documents\github\ProjetoCHOCO\ChocoVerifica.ps1"
Choco_output_file = r"C:\Users\vboxuser\Documents\github\ProjetoCHOCO\chocoOutPut.txt"
comando_choco_install = r"C:\Users\vboxuser\Documents\github\ProjetoCHOCO\ShellExecChoco.ps1"

# Cria o objeto Shell.Application para executar com privilégios
t_shell = win32com.client.Dispatch("Shell.Application")

# Função para executar um script PowerShell elevado
def Verificar_Choco(file_ChocoVerifica):
    t_shell.ShellExecute(
        "powershell.exe",
        f"-ExecutionPolicy Bypass -File \"{file_ChocoVerifica}\"",
        "", 
        "", # runas executa como adm
        0 # uma vez marcado com "0" ele nao abre o prompt só executa o codigo
    )

Verificar_Choco(file_ChocoVerifica)

time.sleep(7)

# 2) Lê o arquivo output.txt, se nele tiver algo, entao o comando foi executando com sucesso
if os.path.exists(Choco_output_file):
    with open(Choco_output_file, "r") as f:
        saida = f.read().strip()
else:
    saida = ""


# verifica se a variavel "saida" esta com o vazia, ou seja, se nela tiver informacao o choco foi instalado!
if not saida: #se tiver vazia instalar o choco
    print("Instalando o CHoco...")

    print("---------------------------------------------------------------------------------------")
    print("Excutando comando shellChoco.ps1 e instalando chocolatey...")
    t_shell = win32com.client.Dispatch("Shell.Application")
    def Instala_Choco(comando_choco_install):
        t_shell.ShellExecute(
        "powershell.exe",
        f"-ExecutionPolicy Bypass -File \"{comando_choco_install}\"",
        "",
        "runas",  # Executa como administrador
        1  # Abre o prompt durante a execução do comando
    )

    
    Instala_Choco(comando_choco_install)

    time.sleep(30)

    Verificar_Choco(file_ChocoVerifica)
    time.sleep(7)

    # 2) Lê o arquivo output.txt, se nele tiver algo, entao o comando foi executando com sucesso
    if os.path.exists(Choco_output_file):
        with open(Choco_output_file, "r") as f:
            saida = f.read().strip()
    else:
        saida = ""
    

else: # se tiver cheia passa direto
    print("O CHoco ja foi instalado!!")



print("--------------------------------------------------------------------------------------------------")
print("---- SE ABAIXO APARECER A VERSO DO CHOCOLATEY, ENTAO ESTA TUDO CERTO SEU CHOCO ESTA INSTALADO ----")
print(f"Choco Instalado com Sucesso {saida}")
print("--------------------------------------------------------------------------------------------------")






