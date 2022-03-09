import xlrd
import typing
import re
# from termcolor import colored

class Controle:

    def __init__(self) -> None:
        self.__planilha_esocial = None
        self.__planilha_candidata = None
        self.__relatorio:typing.IO = None
    

    def set_planilha_esocial(self, nome_pasta_esocial:str) -> None:
        try:
            workbook_esocial = xlrd.open_workbook(nome_pasta_esocial)
            self.__planilha_esocial = workbook_esocial.sheet_by_name('Sheet1')
        except:
            # print("\n\n--- olha o outro erro aí, amigo --- \n\n")
            raise Exception("Falha ao abrir planilha padrão (eSocial).")
        

    def set_planilha_candidata(self, nome_pasta_candidata: str) -> None:
        try:
            workbook_candidata = xlrd.open_workbook(nome_pasta_candidata) 
            self.__planilha_candidata = workbook_candidata.sheet_by_name('Mapeamento Final')
        except:
            raise Exception("Falha ao abrir planilha candidata.")


    def set_relatorio(self, relatorio: typing.IO) -> None:
        self.__relatorio = relatorio

    def escreve_relatorio(self):
        total_de_linhas = min([len(list(self.__planilha_esocial.get_rows())),
                len(list(self.__planilha_candidata.get_rows()))])
        linhas_que_importam = list(range(1, total_de_linhas))
        colunas_que_importam = list(range(0,9))
        espacosExtrasPorCelula = 5

        colunas = "ABCDEFGHIJKLMNOPQRSTUWXYZ"

        def get_string_celula(worksheet_esocial, worksheet_arterh, i:int, j:int, ordinal:int):
            caracteresExibidosPorCelula = [5, 15, 15, 2, 5, 5, 5, 5, 2000]

            conteudo_esocial = ""
            conteudo_arterh = ""

            try:
                conteudo_esocial = int(worksheet_esocial.cell(i, j).value)
            except:
                conteudo_esocial = str(worksheet_esocial.cell(i, j).value)
            finally:
                conteudo_esocial = str(conteudo_esocial)

            conteudo_esocial = re.sub(r"\W", "", conteudo_esocial)[0:caracteresExibidosPorCelula[ordinal]]\
                .ljust(caracteresExibidosPorCelula[ordinal]+espacosExtrasPorCelula, ' ')

            try:
                conteudo_arterh = int(worksheet_arterh.cell(i, j).value)
            except:
                conteudo_arterh = str(worksheet_arterh.cell(i, j).value)
            finally:
                conteudo_arterh = str(conteudo_arterh)

            conteudo_arterh = re.sub(r"\W", "", conteudo_arterh)[0:caracteresExibidosPorCelula[ordinal]]\
                .ljust(caracteresExibidosPorCelula[ordinal]+espacosExtrasPorCelula, ' ')

            return str(conteudo_esocial), str(conteudo_arterh)

        for i in linhas_que_importam:
            for ordinal, j in enumerate(colunas_que_importam):
                conteudo_esocial, conteudo_arterh = get_string_celula(self.__planilha_esocial,
                        self.__planilha_candidata, i, j, ordinal)
                if conteudo_esocial != conteudo_arterh:
                    self.__relatorio.write(f"{colunas[j]}{i+1}")
                self.__relatorio.write("||")
            self.__relatorio.write("\n")

            for ordinal, j in enumerate(colunas_que_importam):
                conteudo_esocial, conteudo_arterh = get_string_celula(self.__planilha_esocial,
                        self.__planilha_candidata, i, j, ordinal)
                if conteudo_esocial != conteudo_arterh:
                    self.__relatorio.write(f"{conteudo_esocial}")
                self.__relatorio.write("||")
            self.__relatorio.write("\n")

            for ordinal, j in enumerate(colunas_que_importam):
                conteudo_esocial, conteudo_arterh = get_string_celula(self.__planilha_esocial,
                        self.__planilha_candidata, i, j, ordinal)
                if conteudo_esocial != conteudo_arterh:
                    self.__relatorio.write(f"{conteudo_arterh}")
                self.__relatorio.write("||")
            self.__relatorio.write("\n")

# Eu não tenho controle sobre o relatório, logo não o fecho neste ponto da execução.
