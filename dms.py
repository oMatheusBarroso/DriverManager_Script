import subprocess
import ctypes
import time
import sys
import os



def Menu():
    print("  ~~~~ Bem-Vindo(a) ao Script Gerenciador de Drivers! ~~~~\
        \n[versão beta de produção] - nov. 2023 - por Matheus Barroso\
        \n-----------------------------------------------------------")

    print("\nO que você deseja fazer?\n")

    print("1 - Aplicação de Drivers")
    print("2 - Extração de Drivers")
    print("3 - Conferir informações do computador")

    print("\n0 - Fechar o Script")

    # Solicita o input, e caso não seja numérico, retorna a mensagem de erro
    try:
        option = int(input("\nDigite o número da sua opção: "))
    except ValueError:
        input("\n  # Erro # -- Insira apenas um número inteiro.\
               \n  Pressione Enter para voltar ao Menu Principal...")
        ResetMenu()

    print("\n-----------------------------------------------------------")

    match option:
        case 1:
            ApplyDrivers()
        case 2:
            ExtractDrivers()
        case 3:
            input("\nFunção Indisponível.\
                   \nPressione Enter para escolher outra opção...")
            ResetMenu()
        case 0:
            sys.exit()
        case _:
            input("\nOpção Indisponível.\
                   \nPressione Enter para escolher outra opção...")
            ResetMenu()

    # Mensagens padrão:

        #input("\nFunção Indisponível. \nPressione Enter para escolher outra opção...")
        #input("\nOpção Indisponível. \nPressione Enter para escolher outra opção...")



# Obtem o modelo da placa mãe do computador
def GetModelName():
    global model
    global displayable_model

    # Grava em temp o resultado do comando wmic, que fornece o modelo
    temp = subprocess.check_output(["wmic", "csproduct", "get", "name"])

    # Nome do modelo para pasta do Windows
    model = temp.decode("utf-8").strip().split("\n")[1].replace(" ", "_").replace("/", "--")
    # Tratamento da saída do comando, removendo todos os caracteres indesejados
        # decode("utf-8") -- decodifica os dados do comando em texto legível
        # strip() -- Remove espaços e linhas saltadas no início e fim da string
        # split("\n")[1] -- Divide a string do modelo pela linha saltada e seleciona a segunda metade, que é o nome em si
        # replace(...) -- Substitui caracteres indesejados em nomes de pastas por outros aceitos pelo Windows

    # Nome do modelo para mostrar ao usuário (pode conter caracteres indesejados em nomes de pastas)
    displayable_model = temp.decode("utf-8").strip().split("\n")[1]



# Obtém o caminho/ diretório desse script
def GetScriptPath():
    global scriptdir

    # Caminho completo do script
    scriptdir = os.path.dirname(os.path.abspath(__file__))
    # .abspath() -- caminho absoluto: inclui o script (...\dms.py)
    # .dirname() -- caminho 'completo': apenas o diretório do caminho absoluto
    # __file__ é uma variável especial em Python que contém o caminho onde o script Python está sendo executado



# Checa a existência das pastas \Driver e \Extrações no diretório desse script
def CheckBasicFolders():
    global drivers_storage
    global extractions_dest

    GetScriptPath()

    # Caminho completo da pasta \Drivers
    drivers_storage = os.path.join(scriptdir, "Drivers")
    # Caminho completo da pasta \Extrações
    extractions_dest = os.path.join(scriptdir, "Extrações")

    # Cria a pasta 'Drivers' no Diretório, caso não exista
    if not os.path.exists(drivers_storage):
        print(f"\n{drivers_storage} será criada!")
        os.mkdir(drivers_storage)

    # Cria a pasta 'Extrações' no Diretório, caso não exista
    if not os.path.exists(extractions_dest):
        print(f"\n{extractions_dest} será criada!")
        os.mkdir(extractions_dest)



# Fornece ao usuário a opção de deletar uma pasta e deleta, se for o caso
def AskAndDeleteFolder(display_path, complete_path):
    x = input(f"\nEncontrei a pasta {display_path}, mas ela está vazia.\
              \nGostaria de deletar a pasta? (S/N): ").upper().strip().replace(" ", "")
    
    if x == "S":
        os.rmdir(complete_path)
        input("Pasta Excluída! Pressione Enter para voltar ao Menu Principal...")
    if x == "N":
        input("Pressione Enter para voltar ao Menu Principal...")

    ResetMenu()



# Função de Extração dos Drivers
def ExtractDrivers():
    global model
    global drivers_storage
    global extractions_dest

    GetModelName()
    CheckBasicFolders()

    # Pasta com nome do modelo em 'Drivers'
    stored_driver = os.path.join(drivers_storage, model)

    # Pasta com nome do modelo em 'Extrações'
    extracted_driver = os.path.join(extractions_dest, model)

    # Comando de extração de drivers (requer privilégios de ADM)
    command = f'DISM /Online /Export-Driver /Destination:"{extracted_driver}"'

    # Checa se a pasta com o nome do modelo já existe em \Drivers
    if os.path.exists(stored_driver):

        # Checa se a pasta está vazia
        if len(os.listdir(stored_driver)) == 0:
            display_path = f"\\Drivers\\{model}\\"
            # Fornece ao usuário opção de deletar a pasta vazia
            AskAndDeleteFolder(display_path, stored_driver)
        
        else:
            print(f"\nJá existem drivers armazenados para este modelo ({displayable_model})!")
            input("Pressione Enter para voltar ao Menu Principal...")
            ResetMenu()

    # Checa se a pasta com o nome do modelo já existe em \Extrações
    if os.path.exists(extracted_driver):
        
        # Checa se a pasta está vazia
        if len(os.listdir(extracted_driver)) == 0:
            display_path = f"\\Extrações\\{model}\\"
            # Fornece ao usuário opção de deletar a pasta vazia
            AskAndDeleteFolder(display_path, extracted_driver)

        else:
            print(f"\nJá existem drivers extraídos para este modelo ({displayable_model})!")
            input("Pressione Enter para voltar ao Menu Principal...")
            ResetMenu()

    # Extrai os drivers caso não existam pastas
    else:

        print(f"\nNão há drivers para este modelo ({displayable_model})!")
        print(f"Os drivers serão salvos em '\\Extrações\\{model}'.")
        confirm = input("Pressione Enter para confirmar a extração...")

        print("\n-----------------------------------------------------------\n")

        # Cria a pasta com nome do modelo em \Extrações
        os.mkdir(extracted_driver)

        print("\nA extração de drivers será executada em outra janela.\
               \nAguarde o fim da sua execução para prosseguir.")
        time.sleep(6) # Tempo de espera para leitura

        # Executa o comando de extração com privilégios de Administrador
        RunAsAdministrator(command)
        
        # O módulo 'subprocess' pode executar como usuário Administrador, mas não com privilégios de ADM
            # subprocess.run(["runas", "/user:Administrator", command], shell=True)
        
        input("\nAo finalizar a execução, pressione Enter...")
        ResetMenu()



# EM CONSTRUÇÃO
# Função de Aplicação dos Drivers
def ApplyDrivers():
    global model
    global drivers_storage

    GetModelName()
    CheckBasicFolders()

    if os.path.exists(os.path.join(drivers_storage, model)):
        print("\nDrivers OK!")
        input("Pressione Enter para voltar ao Menu Principal...")
        ResetMenu()
    else:
        print(f"\nNão há drivers para este modelo ({displayable_model})! :(")
        input("Pressione Enter para voltar ao Menu Principal...")
        ResetMenu()



# Executa um comando com privilégios de Administrador
def RunAsAdministrator(command):
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", f"/K {command}", None, 1)
    except Exception as e:
        print(f"Erro ao executar como administrador: {e}")



# Reseta o Menu -- Limpa o Terminal e executa o Menu de Seleção novamente
def ResetMenu():
    ClearTerminal()
    Menu()



# Limpa o Terminal do Script
def ClearTerminal():
    os.system("cls")



Menu()

# print(f"\nComputador Modelo: {model}")
# print(f"\nDiretório do Script: {scriptdir}")
input("\nOps! :/ Algo de errado não está certo... \nPressione Enter para fechar o Script")