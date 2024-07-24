import tkinter as tk
from tkinter import ttk
import openpyxl
import os


#BACKEND
def load_data():
    diretorio = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(diretorio, 'clientes.xlsx')
    workbook = openpyxl.load_workbook(db_path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)

    for col_name in cols:
        treeview.heading(col_name, text=col_name, anchor='w')

    for value_tuple in list_values[2:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_data():
    nome = nome_entry.get()
    dataNascimento = dataNascimento_entry.get()
    celular = celular_entry.get()
    endereco = endereco_entry.get()
    cidade = cidade_entry.get()
    
    if nome == '' or nome == 'Nome':
        print("Nome é obrigatorio")
    else:
        #Salvar dados no arquivo excel
        diretorio = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(diretorio, 'clientes.xlsx')
        workbook = openpyxl.load_workbook(db_path)
        sheet = workbook.active
        row_values = [nome, dataNascimento, celular, endereco, cidade] #Dados do row
        sheet.append(row_values)
        workbook.save(db_path)

        #Inserir dados salvos na treeview
        treeview.insert('', tk.END, values=row_values)

        #Limpar e colocar os dados iniciais dos nossos campos
        nome_entry.delete(0, 'end')
        nome_entry.insert(0, "Nome")

        dataNascimento_entry.delete(0, 'end')
        dataNascimento_entry.insert(0, "Data de Nascimento")

        celular_entry.delete(0, 'end')
        celular_entry.insert(0, "Celular")

        endereco_entry.delete(0, 'end')
        endereco_entry.insert(0, "Celular")

        cidade_entry.delete(0, 'end')
        cidade_entry.insert(0, "Celular")
        #Limpar e colocar os dados iniciais dos nossos campos
#BACKEND


#FRONTEND
#iniciando a nossa janela
root = tk.Tk()
root.geometry("%dx%d+0+0" % (root.winfo_screenwidth(), root.winfo_screenheight()))
root.attributes('-fullscreen',True)

root.title('Cadastro de Clientes')

#Configuracao do tema
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

#Frame principal
frame = ttk.Frame(root)
frame.pack()
#Frame principal

#Frame do Titulo do app
frame2 = ttk.Frame(frame)
frame2.place(x=50, y=25, width=300, height=100)
text_label = ttk.Label(frame2, text="Cadastro de clientes", font=('Arial', 25))
text_label.pack(expand=True)
#Frame do Titulo do app

#Frame dos Inputs
widgets_frame = ttk.LabelFrame(frame, text="Insira os dados")
widgets_frame.grid(row=0, column=0, padx=50, pady=150)

#INPUTS
nome_entry = ttk.Entry(widgets_frame)
nome_entry.insert(0, "Nome")
nome_entry.bind("<FocusIn>", lambda e: nome_entry.delete('0', 'end'))
nome_entry.grid(row=0, column=0, padx=5, pady=(10,5), ipadx= 100, sticky='ew')

dataNascimento_entry = ttk.Entry(widgets_frame)
dataNascimento_entry.insert(0, "Data de Nascimento")
dataNascimento_entry.bind("<FocusIn>", lambda e: dataNascimento_entry.delete('0', 'end'))
dataNascimento_entry.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

celular_entry = ttk.Entry(widgets_frame)
celular_entry.insert(0, "Celular")
celular_entry.bind("<FocusIn>", lambda e: celular_entry.delete('0', 'end'))
celular_entry.grid(row=2, column=0, padx=5, pady=5, sticky='ew')

endereco_entry = ttk.Entry(widgets_frame)
endereco_entry.insert(0, "Endereço")
endereco_entry.bind("<FocusIn>", lambda e: endereco_entry.delete('0', 'end'))
endereco_entry.grid(row=3, column=0, padx=5, pady=5, sticky='ew')

cidade_entry = ttk.Entry(widgets_frame)
cidade_entry.insert(0, "Cidade")
cidade_entry.bind("<FocusIn>", lambda e: cidade_entry.delete('0', 'end'))
cidade_entry.grid(row=4, column=0, padx=5, pady=5, sticky='ew')

InserirDadosBtn = ttk.Button(widgets_frame, text='Inserir dados'.upper(), command=insert_data)
InserirDadosBtn.grid(row=5, column=0, padx=5, pady=5, sticky='nsew')

sairBtn = ttk.Button(widgets_frame, text='Sair'.upper(), command=root.quit)
sairBtn.grid(row=6, column=0, padx=5, pady=5, sticky='nsew')

#Frame Treeview para visualizar banco de dados
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, padx=50, pady=150)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Nome", "Data de Nascimento", "Celular", "Endereço", "Cidade")
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=35)

treeview.column("Nome", width=120)
treeview.column("Data de Nascimento", width=150)
treeview.column("Celular", width=120)
treeview.column("Endereço", width=120)
treeview.column("Cidade", width=120)
treeview.pack()
treeScroll.config(command=treeview.yview)

load_data()
#Frame Treeview para visualizar banco de dados


#Rodando a nossa janela
root.mainloop()
