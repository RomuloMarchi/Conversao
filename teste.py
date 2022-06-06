from gettext import install

#from extracao_dados import ExtrairDados
import pandas as pd 
from openpyxl import Workbook, load_workbook

from tkinter import messagebox

#from Localizar_arquivos import LocalizarArquivos
#from app import MyGUI
from janela import MyGUI

# class GravarDados:

def DefinirExcel(self,caminho_excel):
    '''
    Define em qual arquivo será gravado os dados do PDF '''
    #Atribuindo uma váriavel a Database 
    exl_path = MyGUI.get_excel(caminho_excel)
    #exl_path = 'abc'
    #arquivo_excel =  MyGUI.get_excel()
    #arquivo_excel = pd.read_excel('C:\\Users\marchrom\Documents\project_build1-main\excel_sample\GSD_Alignment_TT.xlsx')
    #print(exl_path)
    return exl_path
DefinirExcel

janela = MyGUI()
#arquivo = DefinirExcel()
arquivo = janela.get_excel()
print(arquivo)

