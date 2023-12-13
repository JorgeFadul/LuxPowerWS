import ExcelTables
import lxp_wscrap

if __name__ == "__main__":
    lxp_wscrap.DataDownload()
    ExcelTables.excelETL()
