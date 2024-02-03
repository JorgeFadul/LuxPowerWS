import secret
import ExcelTables
import lxp_wscrap
import consumo_prom

user = ""
pword = ""
if __name__ == "__main__":
    try:
        print("LuxPower WebScraping para la Obtención de datos de Corporación Tesla S.A.")

        while not(user==secret.user and pword==secret.password):
            user = str(input("Introduzca su usuario LuxPower: ")).lower()
            pword = str(input("Introduzca su contraseña: "))
            if not(user==secret.user and pword==secret.password):
                print("Usuario o Contraseña incorrecta, intentelo nuevamente")

        lxp_wscrap.DataDownload()
        ExcelTables.excelETL()
        consumo_prom.main_consumo()
        print("Ejecución Terminada: Consulte las carpetas y los archivos creados")
        print("para información detallada de los sistemas LuxPower de Corporacion Tesla S.A.")
        print(f"Carpeta de datos en: {secret.carpeta_descargas}")
        input("Presione la tecla INTRO - ENTER para salir...")

        
    except Exception as e:
        print(f"Se produjo un erorr en la ejecución: {str(e)}")
        input("Presione la tecla INTRO - ENTER para salir...")