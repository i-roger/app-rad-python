import PySimpleGUI as sg # type: ignore

#Inclusions

import os 
import sqlite3

diretorio_corrent = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(diretorio_corrent, 'database.db')

#A tabela será criada se ela não existir dentro do diretório.

connection = sqlite3.connect(db_path)
query = ('''CREATE TABLE IF NOT EXISTS SUPLEMENTO (LOTE CHAR(10), PRODUTO TEXT, FORNECEDOR TEXT)''')
connection.execute(query)

myCursor = connection.cursor()


dados = []
myCursor.execute("select * from SUPLEMENTO")
for i in myCursor:
    dados.append(list(i))

titulos = ['Lote', 'Produto', 'Fornecedor']

layout = [
    [sg.Text(titulos[0]), sg.Input(size=5, key=titulos[0])],
    [sg.Text(titulos[1]), sg.Input(size=20, key=titulos[1])],
    [sg.Text(titulos[2]), sg.Combo(['Fornecedor 1', 'Fornecedor 2', 'Fornecedor 3'], key=titulos[2])],
    [sg.Button('Adicionar'), sg.Button('Editar'), sg.Button('Salvar', disabled=True), sg.Button('Excluir'), sg.Exit('Sair'), sg.Button('Consultar')],
    [sg.Table(dados, titulos, key='tabela')]
]

window = sg.Window('Sistema de gerencia de suplementos', layout)

while True: #Para escutar sempre as ações do usuário!
    
    event, values = window.read()
    print(values)

    if event == 'tabela':
        selected_row = str(values['tabela'][0])
        print(selected_row)

    if event == 'Consultar': #Novo método de exclusão!
        if values['tabela']==[]:
            sg.poup('Nenhuma linha foi selecionada')
        else :
            if sg.popup_ok_cancel('Essa operação não pode ser desfeita. Confirma?') == 'OK':
                try:
                    # INSERT
                    connection = sqlite3.connect(db_path)
                    myCursor = connection.cursor()

                    del dados[values['tabela'][0]]  #Remove a linha selecionada
                    window['tabela'].update(values=dados)
                    PRODUTO = values=dados
                    print("Produto :" + PRODUTO)
                    delete_stmt = "DELETE FROM SUPLEMENTO WHERE PRODUTO ="+PRODUTO
                    myCursor.execute(delete_stmt)
                    # myCursor.execute("DELETE FROM SUPLEMENTO WHERE LOTE = ? ", (values[titulos[0]],))
                    
                    connection.commit()
                    connection.close()

                except sqlite3.Error as e:
                    print(e)

                finally:
                    print("Registro removido!")


    if event == 'Adicionar':
        dados.append([values[titulos[0]], values[titulos[1]],values[titulos[2]]])
        window['tabela'].update(values=dados)
        for i in range(3) : #Limpa as caixas de texto
            window[titulos[i]].update(value='')

        # INSERT
        connection = sqlite3.connect(db_path)
        connection.execute("INSERT INTO SUPLEMENTO (LOTE, PRODUTO, FORNECEDOR) VALUES (?,?,?)", ([values[titulos[0]], values[titulos[1]], values[titulos[2]]]))
        connection.commit()
        connection.close()

    if event == 'Editar' :
        if values['tabela']==[]:
            sg.popup('Nenhuma linha selecionada')
        else:
            editarLinha=values['tabela'][0]
            sg.popup('Editar linha selecionada')
            for i in range(3):
                window[titulos[i]].update(value=dados[editarLinha][i])
            window['Salvar'].update(disabled=False)

    if event == 'Salvar':
        dados[editarLinha]=[values[titulos[0]], values[titulos[1]], values[titulos[2]]]
        window['tabela'].update(values=dados)
        for i in range(3) :
            window[titulos[i]].update(value='')
        window['Salvar'].update(disabled=True)

        # INSERT
        connection = sqlite3.connect(db_path)
        connection.execute("UPDATE SUPLEMENTO set PRODUTO = ?, FORNECEDOR = ? where LOTE = ?", ([values[titulos[0]], values[titulos[1]], values[titulos[2]]]))
        connection.commit()
        connection.close()

    if event == 'Excluir' :
        if values['tabela']==[]:
            sg.poup('Nenhuma linha foi selecionada')
        else :
            if sg.popup_ok_cancel('Essa operação não pode ser desfeita. Confirma?') == 'OK':

                # INSERT
                connection = sqlite3.connect(db_path)
                connection.execute("DELETE FROM SUPLEMENTO WHERE LOTE = ?;", (values[titulos[0]],))
                connection.commit()
                connection.close()

                del dados[values['tabela'][0]]  #Remove a linha selecionada
                window['tabela'].update(values=dados)

    if event in (sg.WIN_CLOSED, 'Sair'):
        break

window.close()