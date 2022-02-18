from cgitb import text
from msilib.schema import File
from tkinter import Tk
from tkinter import ttk
from tkinter import filedialog as fd
from typing import IO

class MainForm (Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Diferenciação de Planilhas")
        self.geometry("400x100")
        self.resizable(False, False)


        self.pasta_padrao: str = ""
        self.lbl_pasta_padrao: ttk.Label = ttk.Label(self, text= "Pasta Padrão: ")
        
        def lbl_pasta_padrao_click() -> None:
            self.pasta_padrao = fd.askopenfilename(filetypes=(('planilha xlsx', '*.xlsx'),))
            self.lbl_pasta_padrao.config(text="Pasta Padrão: " + self.pasta_padrao)
            if valida(self.pasta_padrao, self.pasta_candidata, self.relatorio):
                executa(self.lbl_status)

        self.lbl_pasta_padrao.grid()
        self.lbl_pasta_padrao.bind("<Button-1>", lambda e: lbl_pasta_padrao_click())


        self.pasta_candidata: str = ""
        self.lbl_pasta_candidata: ttk.Label = ttk.Label(self, text= "Pasta Candidata: ")
        
        def lbl_pasta_candidata_click() -> None:
            self.pasta_candidata = fd.askopenfilename(filetypes=(('planilha xlsx', '*.xlsx'),))
            self.lbl_pasta_candidata.config(text="Pasta Candidata: " + self.pasta_candidata)
            if valida(self.pasta_padrao, self.pasta_candidata, self.relatorio):
                executa(self.lbl_status)

        self.lbl_pasta_candidata.grid()
        self.lbl_pasta_candidata.bind("<Button-1>", lambda e: lbl_pasta_candidata_click())


        self.relatorio: IO
        self.lbl_relatorio: ttk.Label = ttk.Label(self, text= "Relatório: ")
        
        def lbl_relatorio_click() -> None:
            try:
                with fd.asksaveasfile(filetypes=(('text file', '*.txt'),)) as self.relatorio:
                    self.lbl_relatorio.config(text="Relatorio: " + self.relatorio.name)
                    if valida(self.pasta_padrao, self.pasta_candidata, self.relatorio):
                        executa(self.lbl_status)
            finally:
                pass

        self.lbl_relatorio.grid()
        self.lbl_relatorio.bind("<Button-1>", lambda e: lbl_relatorio_click())


        self.status = ""
        self.lbl_status: ttk.Label = ttk.Label(self, text= "Status: ")
        
        self.lbl_status.grid()

        def valida (plan_padrao: str, plan_candidata: str, relatorio: str) -> bool:
            return plan_padrao != "" and plan_candidata != "" and relatorio != ""

        def executa(status: ttk.Label) -> None:
            status.config(text="Status: " + "Modificado!")
            status.config(foreground='blue')



if __name__ == '__main__':
    MainForm().mainloop()
