from gettext import install

#from extracao_dados import ExtrairDados
import pandas as pd 
from openpyxl import Workbook, load_workbook

from tkinter import messagebox
#from Localizar_arquivos import LocalizarArquivos

class GravarDados:

    def DefinirExcel():
        '''
        Define em qual arquivo será gravado os dados do PDF '''
        #Atribuindo uma váriavel a Database 

        #arquivo_excel = LocalizarArquivos.get_excel()
        arquivo_excel = pd.read_excel('C:\\Users\marchrom\Documents\project_build1-main\excel_sample\GSD_Alignment_TT.xlsx')
        print(arquivo_excel)

        # #Atribuindo uma váriavel a coluna de Códigos
        codigos_excel = arquivo_excel.iloc[:,0]
        
        # #Atribuindo uma váriavel a coluna de Descrições
        descricao_excel = arquivo_excel.iloc[:,1]

        # #Atribuindo uma váriavel a coluna de SMVs
        smv_excel = arquivo_excel.iloc[:,2]

        print(codigos_excel,descricao_excel,smv_excel)
        return descricao_excel, smv_excel, codigos_excel
     
    DefinirExcel()
 

    #Definindo uma variavel para a planilha do arquivo excel        
    #planilha = load_workbook(DefinirExcel().arquivo_excel)
    planilha = load_workbook('C:\\Users\marchrom\Documents\project_build1-main\excel_sample\GSD_Alignment_TT.xlsx')

    #Definindo a aba ativa do Excel será o local de gravação dos dados(É sempre definido como a planilha visualizada por útlimo)
    aba_ativa = planilha.active



    def GravarDados(self,lista_de_codigos, lista_descricoes, lista_de_smv, codigos_excel, descricao_excel, smv_excel, aba_ativa):
        a = 0
        b = 0
        c = 0


        for codigos in len(codigos_excel):
            aba_ativa[f"A{codigos}"] = lista_de_codigos[a]
            codigos += 1
        
        for descricoes in len(descricao_excel):
            aba_ativa[f"B{descricoes}"] = lista_descricoes[b]
            descricoes += 1

    
        for smv in len(smv_excel):
            aba_ativa[f"C{smv}"] = lista_de_smv[c]
            smv += 1
    
    messagebox.showinfo(tittle=None, message= 'Conversão realizada!')

    
    #Grava e substitui o arquivo selecionado como DB.
    planilha.save('arquivo.xlsx')

    # print(srcCod)
    # print(srcDesc)
    # print(srcSmv)
    # print(pCod)

    
