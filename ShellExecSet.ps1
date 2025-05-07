& {
    'LocalMachine','CurrentUser' |
    ForEach-Object {
        Set-ExecutionPolicy AllSigned -Scope $_ -Force -Verbose
    }
} 2>$null 3>$null 4>&1 |
  Out-File -FilePath 'C:\Users\vboxuser\Documents\github\ProjetoCHOCO\output.txt' -Encoding UTF8
