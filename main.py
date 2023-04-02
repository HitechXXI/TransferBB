import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import minhasdefs
import os
from tkinter import filedialog
import wx.grid

""" COMANDO PARA GERAR O EXE: pyinstaller -w --hidden-import babel.numbers main.py"""

# =======================================================================

municipios_selecionados = []

# =======================================================================

def obtertransferencias():

    colunas = []
    colunas_inicial = minhasdefs.tabulartransferencias()
    colunas.append(colunas_inicial[0])

    colunas_inicial.pop(0)
    numerocolunas_inicial = len(colunas_inicial)
    arquivo = colunas_inicial[numerocolunas_inicial - 1]
    colunas_inicial = colunas_inicial[:numerocolunas_inicial - 1]

    colunas_inicial.sort(reverse=False)

    # SOLICITA QUE O USUÁRIO ESPECIFIQUE UM NOME PARA SALVAR O ARQUIVO DE RESULTADO
    # ----------------------------------------------------------------------------------------
    files = [('Text Document', '*.txt')]
    arquivofinal = filedialog.asksaveasfilename(filetypes=files, defaultextension=files)
    arqnome_aux = arquivofinal

    try:
        os.remove(arquivofinal)
    except:
        pass

    arquivofinal = arqnome_aux
    # ----------------------------------------------------------------------------------------

    for elemento in colunas_inicial:
        colunas.append(elemento)

    numerocolunas = len(colunas) + 1

    for col in range(0, numerocolunas - 1):
        if col != numerocolunas - 2:
            open(arquivofinal, "a", encoding="utf-8").write(colunas[col] + "\t")
        else:
            open(arquivofinal, "a", encoding="utf-8").write(colunas[col] + "\n")

    with open(arquivo, 'r', encoding="utf-8") as fp:
        for line in fp:
            index = line.find(" - RN")
            index_aux = index - 1
            while line[index_aux] != '\t':
                index_aux -= 1
            index_aux += 1
            municipio = line[index_aux: index]

            transferencias = []

            #transferencias.append(municipio)

            # A variável "colunas" é uma lista
            for coluna in colunas[1:numerocolunas - 1]:
                if line.find(coluna) != -1:
                    index = line.find(coluna)
                    index = line.find("CREDITO FUNDO	R$ ", index, line.find('\n'))
                    index += 17
                    index_aux = index
                    while line[index_aux] != "C":
                        index_aux += 1
                    #transferencias.append({'repasse': coluna, 'valor': line[index:index_aux - 1]})
                    transferencias.append({coluna: line[index:index_aux - 1]})
                else:
                    #transferencias.append({'repasse': coluna, 'valor': '0,00'})
                    transferencias.append({coluna: '0,00'})
                index += 1

            #transferencias.sort(reverse=False, key=lambda element: element['repasse'])
            transferencias = [municipio, *transferencias]
            #print(transferencias)

            open(arquivofinal, "a", encoding="utf-8").write(transferencias[0] + "\t")
            for col in range(1, numerocolunas - 1):
                if col != numerocolunas - 2:
                    #print(transferencias[col])
                    # A linha abaixo imprimi no arquivo texto um tab após o valor de cada repasse
                    open(arquivofinal, "a", encoding="utf-8").write(transferencias[col][colunas[col]] + "\t")
                else:
                    # A linha abaixo imprimi pula para a próxima linha do arquivo
                    open(arquivofinal, "a", encoding="utf-8").write(transferencias[col][colunas[col]] + "\n")
    fp.close()

    messagebox.showinfo("Tranferência BB", "Finish! Transferências obtidas com sucesso, gravados no arquivo " +
                                        arquivofinal)

# =======================================================================

def pegardados():

    municipios_selecionados = []

    """#files = [('All Files', '*.*'), ('Python Files', '*.py'), ('Text Document', '*.txt')]
    files = [('Text Document', '*.txt')]
    arqnome = filedialog.asksaveasfilename(filetypes = files, defaultextension = files)
    arqnome_aux = arqnome

    try:
        os.remove(arqnome)
    except:
        pass

    arqnome = arqnome_aux"""

    for i in listbox_selecionarmunicipio.curselection():
        municipios_selecionados.append(listbox_selecionarmunicipio.get(i).replace('\n', ""))
    if len(municipios_selecionados) == 0 or data_inicial.get() == "" or data_final.get() == "" or data_inicial.get() > data_final.get():
        messagebox.showinfo("Tranferência BB", "Verifique se selecionou pelo menos um município, se prencheu as datas corretamente.")
    else:
        minhasdefs.obtertransferenciabb(municipios_selecionados, combobox_selecionaruf.get(), data_inicial.get(),
                                    data_final.get())

# =======================================================================

label_tarefa_texto = "Copyright: João Evangelista"

janela = tk.Tk()

janela.title("Ferramenta de obtenção das transferências do BB Arrecadação")
# O redimensionamento não automático implica 'weight=0'
janela.rowconfigure(0, weight=1)
janela.columnconfigure([0, 1], weight=1)
janela.config(padx=10)

label_uf = tk.Label(text="UF:")
label_uf.grid(row=0, column=0, padx=5, pady=10, sticky='nsew')
lista_uf = minhasdefs.listaruf()
combobox_selecionaruf = ttk.Combobox(values=lista_uf, state='readonly')
combobox_selecionaruf.current(0)
combobox_selecionaruf.grid(row=0, column=1, pady=10, sticky='nsew')

label_municipio = tk.Label(text="Município:")
label_municipio.grid(row=1, column=0, padx=5, pady=10, sticky='nsew')
lista_municipios = minhasdefs.listarmunicipio()
lista_municipios_var = tk.StringVar(value=lista_municipios)
listbox_selecionarmunicipio = tk.Listbox(janela, listvariable=lista_municipios_var, height=6, selectmode='extended')
listbox_selecionarmunicipio.grid(row=1, column=1, sticky='nsew')

# link a scrollbar to a list
scrollbar = ttk.Scrollbar(janela, orient='vertical', command=listbox_selecionarmunicipio.yview)
listbox_selecionarmunicipio['yscrollcommand'] = scrollbar.set
scrollbar.grid(row=1, column=2, sticky='ns')

label_datainicial = tk.Label(text="Data Inicial:")
label_datainicial.grid(row=4, column=0, padx=5, pady=10, sticky='nsew')
data_inicial = tk.StringVar()
data_inicial = DateEntry(janela, selectmode='day', date_pattern='dd/mm/yyyy', textvariable=data_inicial)
data_inicial.grid(row=4, column=1, pady=10, sticky='nsew')

label_datafinal = tk.Label(text="Data Final:")
label_datafinal.grid(row=6, column=0, padx=5, pady=10, sticky='nsew')
data_final = tk.StringVar()
data_final = DateEntry(janela, selectmode='day', date_pattern='dd/mm/yyyy', textvariable=data_final)
data_final.grid(row=6, column=1, pady=10, sticky='nsew')

botao_pegardados = tk.Button(text="Pegar Dados", command=pegardados)
botao_pegardados.grid(row=8, column=0, padx=10, pady=10, columnspan=4, sticky='nsew')

botao_gerarxlsx = tk.Button(text="Gerar arquivo txt", command=obtertransferencias)
botao_gerarxlsx.grid(row=9, column=0, padx=10, pady=10, columnspan=4, sticky='nsew')

label_tarefa = tk.Label(text=label_tarefa_texto)
label_tarefa.grid(row=10, column=0, padx=10, pady=10, columnspan=4, sticky='nsew')

janela.mainloop()

