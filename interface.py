from tkinter import *
import relatorio_efd
import time

class App:
    def __init__(self,master=None):

        self.container = Frame(master)
        self.container["padx"] = 80
        self.container["pady"] = 5
        self.container.pack()
        self.container.configure(bg='#333333')

        self.label_file_explorer = Label(self.container,
                                    text="Relatório de Entradas SPED EFD",
                                    width=30, height=4,bg="#333333",
                                    fg="white")
        self.label_file_explorer["font"] =("Calibri", "11", "bold")

        self.label1 = Label(self.container,bg="#333333", fg="white", width=15,height=-1)
        self.button_exportar = Button(self.container,
                             text="Escolher arquivo SPED EFD",bg="#4c4c4c",fg="white",width=25,
                             command=lambda:[relatorio_efd.Sped().exportar_excel(),self.fdb_3()])
        self.button_exportar["font"] = ("Calibri", "10", "bold")

        self.label_file_explorer.grid(column=1, row=1)
        self.label1.grid(column=1, row=3)
        self.button_exportar.grid(column=1, row=4)
        self.label2 = Label(self.container, text="", bg="#333333", fg="white",
                            width=30, pady=20, padx=10)
        self.label2.grid(column=1, row=5, )



    def fdb_1(self):
        self.label2 = Label(self.container,text="Aguarde enquanto arquivo é importado ...", bg="#333333", fg="white", width=30,pady=20,padx=10 )
        self.label2.grid(column=1, row=5, )
    def fdb_2(self):
        self.label2 = Label(self.container,text="Arquivo importado", bg="#333333", fg="white", width=30, )
        self.label2.grid(column=1, row=5, )
    def fdb_3(self):
        self.label2 = Label(self.container,text="Arquivo exportado com sucesso", bg="#333333", fg="white", width=30, )
        self.label2.grid(column=1, row=5, )

if __name__ == "__main__":
    root = Tk()
    frame = Frame()
    root.title('SPED to EXCEL')
    root.configure(bg='#333333')
    frame.pack(expand=True, fill=BOTH)
    App(root)
    root.mainloop()
