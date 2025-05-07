import time
import os
import win32com.client

# Caminhos dos scripts PowerShell e do arquivo de saída
comando_set = r"C:\Users\vboxuser\Documents\github\ProjetoCHOCO\ShellExecSet.ps1"
output_file = r"C:\Users\vboxuser\Documents\github\ProjetoCHOCO\output.txt"

# Cria o objeto Shell.Application para executar com privilégios
t_shell = win32com.client.Dispatch("Shell.Application")


# Função para executar um script PowerShell elevado
def executar_powershell_alteracao_politicas(script_path):
    t_shell.ShellExecute(
        "powershell.exe",
        f"-ExecutionPolicy Bypass -File \"{script_path}\"",
        "",
        "runas",
        0
    )


# Seta a politica para Allsigned e armazena a resposta num arquivo output.txt
print(f"Alterando a Política atual para AllSigned via ShellExecSet.ps1...")
# Executa o script que ajusta a política para AllSigned
executar_powershell_alteracao_politicas(comando_set)
time.sleep(8)

# 2) Lê o arquivo output.txt, se nele tiver allsigned, entao o comando foi executando com sucesso
if os.path.exists(output_file):
    with open(output_file, "r") as f:
        saida = f.read().strip()
else:
    saida = ""


print("--------------------------------------------------------------------------------------------------")
# confirma se as politicas estao realmente
print(f"Política atual: {saida}")


