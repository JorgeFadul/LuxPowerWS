import ExcelTables
import lxp_wscrap
import consumo_prom

if __name__ == "__main__":
    lxp_wscrap.DataDownload()
    ExcelTables.excelETL()
    consumo_prom.main_consumo()
    print("Ejecución Terminada: Consulte las carpetas y los archivos creados\npara información detallada de los sistemas LuxPower")