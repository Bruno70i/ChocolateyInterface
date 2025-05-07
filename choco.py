import subprocess
import time
import ctypes






# Comando para verificar a politica do Windoes
comando = r"Get-ExecutionPolicy"

# Executa o comando da variavel "comando"  e armazena essa informaçao na variavel "resultado"
resultado = subprocess.run(["powershell", "-Command", comando,],
    capture_output=True,
    text=True
)

# Exibe a saída do comando
print(resultado.stdout)

# Exibe a saída do comando
print(resultado.stderr)




if resultado == "AllSigned":
    # para o codigo
    print("codigo parou")
 
    
else:
    NovoComando = r"Set-ExecutionPolicy AllSigned -Scope CurrentUser"

    resultado = subprocess.run(["powershell", "-Command", NovoComando,],
    capture_output=True,
    text=True
    )

    print("politica alterada")

    # Exibe a saída do comando
    print(resultado.stdout)

    # Exibe a saída do comando
    print(resultado.stderr)























