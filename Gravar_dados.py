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
        
     
    DefinirExcel()
 

    #Definindo uma variavel para a planilha do arquivo excel        
    #planilha = load_workbook(DefinirExcel().arquivo_excel)
    planilha = load_workbook('C:\\Users\marchrom\Documents\project_build1-main\excel_sample\GSD_Alignment_TT.xlsx')

    #Definindo a aba ativa do Excel será o local de gravação dos dados(É sempre definido como a planilha visualizada por útlimo)
    aba_ativa = planilha.active



    #Usando openpyplx para gravar os dados na planilha selecionada

    #Definindo váriaveis Source das listas geradas do PDF
    #Atribuindo uma váriavel a coluna de Códigos
    # srcCod = lista_de_codigos
    #Atribuindo uma váriavel a coluna de Descrições
    # srcDesc = lista_descricoes
    #Atribuindo uma váriavel a coluna de SMVs
    # srcSmv = lista_de_smv


    #Criando ponteiros para rodar os loopings
    # celulaA = 0
    # celulaB = 0
    # celulaC = 0
    # a = 0
    # b = 0
    # c = 0

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

    




    #Atribuindo uma váriavel a quantidade de linhas da Código (Usando como ponteiro nos for)
    # pCod = len(dbExcel.iloc[:,0]) 

    #Escreve na próxima célula em branco da coluna A(apontada por pCod) até o tamanho de srcCod

    # for celulaA in srcCod:
    #     aba_ativa[f"A{pCod}"] = srcCod[a]
    #     pCod = pCod + 1
    #     a = a + 1

    
    #Atribuindo uma váriavel a quantidade de linhas da Descrição (Usando como ponteiro nos for)
    # pDesc = len(dbExcel.iloc[:,0]) 

    #Escreve na próxima célula em branco da coluna A(apontada por pDesc) até o tamanho de srcCod

    # for celulaB in srcDesc:
    #     aba_ativa[f"B{pDesc}"] = srcDesc[b]
    #     pDesc = pDesc + 1
    #     b = b + 1

    #Atribuindo uma váriavel a quantidade de linhas da SMV (Usando como ponteiro nos for)
    # pSmv = len(dbExcel.iloc[:,1])


    #Escreve na próxima célula em branco da coluna A(apontada por pDesc) até o tamanho de srcCod

    # for celulaC in srcSmv:
    #     aba_ativa[f"C{pSmv}"] = srcSmv[c]
    #     pSmv = pSmv + 1
    #     c = c + 1


    #Grava e substitui o arquivo selecionado como DB.
    planilha.save('arquivo.xlsx')

    # print(srcCod)
    # print(srcDesc)
    # print(srcSmv)
    # print(pCod)

    
