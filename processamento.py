from Gravar_dados import GravarDados
from extracao_dados import ExtrairDados
#from util import Util
#from app import MyGUI



class ProcessamentoDeDados:
    def processamento(self, pdf_path, excel_path):
        
        # # #lista de dados extraidos
        # lista_dados_extraidos = ExtrairDados().extrair_dados(pdf_path)

        # # #lista de códigos extraidos
        # lista_de_codigos = ExtrairDados().extrair_lista_codigos(lista_dados_extraidos, pdf_path)

        # # #lista de descricoes do PDF
        # lista_descricoes = ExtrairDados().extrair_descricao(lista_dados_extraidos, lista_de_codigos)

        # # #extracao de SMV
        # lista_de_smv = ExtrairDados().extrair_smv(lista_dados_extraidos, lista_de_codigos)


       
        excel = GravarDados()
        excel.DefinirExcel(excel_path)
        #gravação dos dados no excel
        #excel.GravarDados()

        #Util().escrever_dados_planilha(lista_de_codigos, lista_descricoes, lista_de_smv)

        pass






