from tkinter import filedialog
from tkinter import messagebox


class LocalizarArquivos:

    def buscar_pdf(self, label_pdf):
        """
        solicita ao usuario localizar o pdf através do explorer
        """
        caminho_do_arquivo_pdf = filedialog.askopenfilename(initialdir="/",
                                                      title="Selecionar o Arquivo",
                                                      filetypes=(('Adobe Acrobat Document', '*.pdf'),
                                                                 ("all files", "*.*")))

        label_pdf.configure(text="Pdf Selecionado!")
        
        return caminho_do_arquivo_pdf

    def buscar_excel(self, label_excel):
        """
        solicita ao usuario localizar o excel através do explorer
        """
        caminho_do_arquivo_excel = filedialog.askopenfilename(initialdir="/",
                                                 title="Selecionar o Arquivo",
                                                 filetypes=(('Microsoft Excel Worksheet', '*.xlsx'),
                                                            ("all files", "*.*")))
        label_excel.configure(text=caminho_do_arquivo_excel)
        
        return caminho_do_arquivo_excel