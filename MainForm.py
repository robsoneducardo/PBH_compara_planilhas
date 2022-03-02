from cgitb import text
from msilib.schema import File
from tkinter import Tk
from tkinter import ttk
from tkinter import filedialog as fd
from typing import IO

from Controle import Controle

class MainForm (Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Diferenciação de Planilhas")
        self.geometry("600x100")
        self.resizable(False, False)
        
        self.__controle: Controle = Controle()

        self.__pasta_padrao: str = ""
        self.__pasta_candidata: str = ""
        self.__relatorio: str = ""


        self.lbl_pasta_padrao: ttk.Label = ttk.Label(self, text= "Pasta Padrão: ")
        
        def lbl_pasta_padrao_click() -> None:
            self.__pasta_padrao = fd.askopenfilename(filetypes=(('planilha xls', '*.xls'),))
            try:
                self.__controle.set_planilha_esocial(self.__pasta_padrao)
                self.lbl_pasta_padrao.config(text="Pasta Padrão: " + self.__pasta_padrao)
                if valida(self.__pasta_padrao, self.__pasta_candidata, self.__relatorio):
                    self.lbl_status.config(text="Gravando relatório...")
                    self.__controle.escreve_relatorio()
                    self.lbl_status.config(text="Relatório gravado.")
            except Exception as e:
                self.lbl_status.config(text="Erro: " + str(e))

        self.lbl_pasta_padrao.grid()
        self.lbl_pasta_padrao.bind("<Button-1>", lambda e: lbl_pasta_padrao_click())


        self.lbl_pasta_candidata: ttk.Label = ttk.Label(self, text= "Pasta Candidata: ")
        
        def lbl_pasta_candidata_click() -> None:
            self.__pasta_candidata = fd.askopenfilename(filetypes=(('planilha xls', '*.xls'),))
            try:
                self.__controle.set_planilha_candidata(self.__pasta_candidata)
                self.lbl_pasta_candidata.config(text="Pasta Candidata: " + self.__pasta_candidata)
                if valida(self.__pasta_padrao, self.__pasta_candidata, self.__relatorio):
                    self.lbl_status.config(text="Gravando relatório...")
                    self.__controle.escreve_relatorio()
                    self.lbl_status.config(text="Relatório gravado.")
                else:
                    self.lbl_status.config(text="Não validou.")
            except Exception as e:
                self.lbl_status.config(text="Erro: " + str(e))

        self.lbl_pasta_candidata.grid()
        self.lbl_pasta_candidata.bind("<Button-1>", lambda e: lbl_pasta_candidata_click())


        self.lbl_relatorio: ttk.Label = ttk.Label(self, text= "Relatório: ")
        
        def lbl_relatorio_click() -> None:
            try:
                with fd.asksaveasfile(filetypes=(('text file', '*.txt'),)) as relatorio:
                    self.__controle.set_relatorio(relatorio)
                    self.lbl_relatorio.config(text="Relatorio: " + relatorio.name)
                    self.__relatorio = relatorio.name
                    if valida(self.__pasta_padrao, self.__pasta_candidata, self.__relatorio):
                        self.lbl_status.config(text="Gravando relatório...")
                        self.__controle.escreve_relatorio()
                        self.lbl_status.config(text="Relatório gravado.")
            except Exception as e:
                self.lbl_status.config(text="Erro: " + str(e))

        self.lbl_relatorio.grid()
        self.lbl_relatorio.bind("<Button-1>", lambda e: lbl_relatorio_click())


        self.lbl_status: ttk.Label = ttk.Label(self, text= "Status: ")
        self.lbl_status.grid()

        def valida (plan_padrao: str, plan_candidata: str, relatorio: str) -> bool:
            return plan_padrao != "" and plan_candidata != "" and relatorio != ""



if __name__ == '__main__':
    MainForm().mainloop()
