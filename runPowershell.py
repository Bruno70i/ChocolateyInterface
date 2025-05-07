import win32com.shell.shell as shell

# Caminho do script PowerShell
ps_script = r"D:\github\ProjetoCHOCO\getproc.ps1"

# Executa o PowerShell como administrador
shell.ShellExecuteEx(
    lpVerb="runas",  # Solicita elevação
    lpFile="powershell.exe",
    lpParameters=f"-ExecutionPolicy Bypass -File \"{ps_script}\"",
    nShow=1  # Mostra a janela
)