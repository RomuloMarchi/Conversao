import tkinter as tk
from tkinter import *
from tkinter import filedialog
#from Gravar_dados import GravarDados

from Localizar_arquivos import LocalizarArquivos
from processamento import ProcessamentoDeDados

class MyGUI(tk.Frame):
    def  __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        global browse_pdf_btn
        global browse_excels_btn
        canvas = tk.Canvas(self, width=720, height=460, relief="raised")
        canvas.pack()

        canvas.create_image( 0, 0, 
                     anchor = "nw") # image = bg,

        browse_pdf_btn = tk.Button(text="Selecione o arquivo PDF ",
                                      bg='Crimson', fg='white', font=('helvetica', 10, 'bold'),
                                      command=self.get_pdf)
        canvas.create_window(380, 80, window=browse_pdf_btn)


        browse_excels_btn = tk.Button(text="Selecione o arquivo Excel ",
                                      bg='DarkCyan', fg='white', font=('helvetica', 10, 'bold'),
                                      command=self.get_excel)
        canvas.create_window(380, 120, window=browse_excels_btn)

        browse_confirm_btn = tk.Button(text="Iniciar Migracão de Arquivos",
                                      bg='MediumSlateBlue', fg='white', font=('helvetica', 10, 'bold'),
                                      command=self.launch_process)
        canvas.create_window(380, 250, window=browse_confirm_btn)

    def get_excel(self):
        global caminho_excel
        print('comando excel')
        caminho_excel = LocalizarArquivos().buscar_excel(browse_excels_btn)
        print(caminho_excel)
        return caminho_excel
   
    def get_pdf(self):
        global caminho_pdf
        print('comando pdf')
        caminho_pdf = LocalizarArquivos().buscar_pdf(browse_pdf_btn)
        return caminho_pdf

    def launch_process(self):       
        ProcessamentoDeDados().processamento(pdf_path=caminho_pdf, excel_path = caminho_excel)
        
        print('dados processados')
        



if __name__ == "__main__":
    
    root = tk.Tk()
    root.title('Adidas Tool')
    root.wm_geometry("720x460")
    bg = PhotoImage(file = "C:\\Users\marchrom\Documents\Imagens\logo_adidas.png")
    main = MyGUI(root)
    main.pack(side="top", fill="both", expand=True)
    root.resizable(False,False)
    root.mainloop()