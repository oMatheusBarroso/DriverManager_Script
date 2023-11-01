import subprocess

print("~ Bem-Vindo(a) ao Script Gerenciador de Drivers! ~")

def GetModelName():
    global model
    temp = subprocess.check_output(["wmic", "csproduct", "get", "name"])
    model = temp.decode("utf-8").strip().split("\n")[1].replace(" ", "_").replace("/","-")

#def CreateFolder():
    #global model

#def ExtractDrivers():

#def ApplyDrivers():

GetModelName()

print("\nComputador Modelo: " +model)
input("\nPressione Enter para fechar o Script")