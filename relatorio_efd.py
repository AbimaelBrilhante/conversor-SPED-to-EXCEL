import csv
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename,asksaveasfile

class Sped:
    def __init__(self):
        pass

    def tratamento_arquivo(self):
        Tk().withdraw()
        arquivo_sped = askopenfilename()
        with open(arquivo_sped, 'r', encoding="ANSI") as file:
            reader = csv.reader(file, delimiter='|')
            self.rows = list(reader)
            self.indicec100 = []
            self.indicec170 = []
            self.efd = []

            # FILTRAGEM DOS REGISTROS C100 (DE ENTRADA) E C170 DO ARQUIVO
            for efd_filt in self.rows:
                if  efd_filt != [] and efd_filt[0] == "":
                    if (efd_filt[1] == "C100" and efd_filt[2] == "0" and efd_filt[3] == "1") or efd_filt[1] == "C170":
                        self.efd.append(efd_filt)

            # NUMERO DA LINHA QUE SE ENCONTRA O REGISTRO C100
            for indice, reg100 in enumerate(self.efd):
                if reg100[1] == "C100":
                    self.indicec100.append(indice)

            # NUMERO DA LINHA QUE SE ENCONTRA O REGISTRO C170
            for h, reg170 in enumerate(self.efd):
                if reg170[1] == "C170":
                    self.indicec170.append(h)

            # FAZENDO APPEND DO REGISTRO C100 E C170 PARA DENTRO DE UMA LISTA (PARA CADA C170 O C100 DEVERÁ SE REPETIR)
        x = 0
        y = 1
        self.contador = 0
        self.relatorio = []

        while (self.contador < len(self.indicec100) - 1):
            for r in range(self.indicec100[x] + 1, self.indicec100[y]):
                self.relatorio.append(self.efd[self.indicec100[x]] + self.efd[r])
            x += 1
            y += 1
            self.contador += 1

            # IDENTIFICANDO O ULTIMO C100 E SEUS CORRESPONDENTES (C170) E FAZENDO O APPEND NA MESMA LISTA DE CIMA

        quantidade_dos_ultimos_c170 = self.indicec170[-1] - self.indicec100[-1]
        k = 0
        while k < quantidade_dos_ultimos_c170:
            self.relatorio.append(
                self.efd[self.indicec100[len(self.indicec100) - 1]] + self.efd[self.indicec170[len(self.indicec170) - quantidade_dos_ultimos_c170 + k]])
            k += 1
        return self.relatorio

    def exportar_excel(self):
        self.tratamento_arquivo()
        df = pd.DataFrame(self.relatorio)
        #df.to_csv('arquivo salvo.csv', sep=';')
        files = [('CSV', '*.csv')]
        df.to_csv(asksaveasfile(mode="w",filetypes=files, defaultextension=files), header=False, index=False, sep=';')



if __name__ == "__main__":


    Sped().exportar_excel()









#FAZER EXPORTAÇÃO PARA O EXCEL DIRETAMENTE DO ARQUIVO TXT
#REVERTER DO EXCEL PARA O TXT
#exportar outros registros sped


