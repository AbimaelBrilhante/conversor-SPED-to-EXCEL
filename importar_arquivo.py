from tkinter import *
import main

class App:
    def __init__(self):

        self.window = Tk()
        self.window.title('SPED to Excel')
        self.window.geometry("250x200")
        self.window.config(background="#333333")
        self.label_file_explorer = Label(self.window,
                                    text="Relatório de Entradas SPED EFD",
                                    width=30, height=4,bg="#333333",
                                    fg="white")
        self.label_file_explorer["font"] =("Calibri", "11", "bold")
        self.button_explore = Button(self.window,
                                text="Browse Files",bg="#4c4c4c",fg="white",width=15,
                                command=lambda:[print("aguarde"),(main.Sped('sped_relatorio.db').adicionar_bd()),print("sucesso")])
        self.button_explore["font"] = ("Calibri", "10", "bold")
        self.label1 = Label(self.window,bg="#333333", fg="white", width=15,height=-1)
        self.button_exportar = Button(self.window,
                             text="Exportar para excel",bg="#4c4c4c",fg="white",width=15,
                             command=main.Sped('sped_relatorio.db').exportar_excel)
        self.button_exportar["font"] = ("Calibri", "10", "bold")
        self.label2 = Label(self.window, bg="#333333", fg="white", width=30, )
        self.label_file_explorer.grid(column=1, row=1)
        self.button_explore.grid(column=1, row=2)
        self.label1.grid(column=1, row=3)
        self.button_exportar.grid(column=1, row=4)
        self.label2.grid(column=1, row=3,)
        self.window.mainloop()

App()
