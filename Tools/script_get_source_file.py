import os
import tkinter
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox as mbox


class GetDadosGUI:


    def __int__(self):
        self.source_planilha = ""

    def main(self):

        def upload_file():
            conteudo = tkinter.filedialog.askopenfilename(title='Selecione a planilha')
            self.source_planilha = conteudo
            print(f"- [Planilha]: {self.source_planilha}")
            return app.destroy()

        def alterar_cor_button(button, cor_ativa, cor_apagada):
            button.bind("<Enter>", func=lambda e: button.config(foreground=cor_ativa))
            button.bind("<Leave>", func=lambda e: button.config(foreground=cor_apagada))

        img = fr"{os.getcwd()}\Tools\sistema.png"
        app = tkinter.Tk()

        C = tkinter.Canvas(app, height=200, width=300)
        app.title('Robô-PJeCalc')
        filename = tkinter.PhotoImage(file=img)
        background_label = tkinter.Label(app, image=filename)

        background_label.place(x=0, y=0, relwidth=1, relheight=1)
        btn = tkinter.Button(app, text='Clique Aqui', font=('Helvetica', '8'), foreground='#28506E', command=upload_file,
                             borderwidth=0, background='#FFFFFF', activeforeground='green', activebackground='#FFFFFF')
        btn.place(height=20, width=120, x=90, y=148)

        alterar_cor_button(btn, "red", "#28506E")

        C.pack()
        app.mainloop()