from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui
import pyperclip
import os
from tkinter import filedialog
from tkinter import messagebox
# =======================================================================

def listaruf():
    listauf = ['RN']
    return listauf

# =======================================================================

def listarmunicipio():
    lista_municipios = []
    try:
        # Lê o dada frame de municípios e anexa à lista "lista_municípios"
        arquivo = "municipiosRN.txt"
        with open(arquivo, 'r', encoding="utf-8") as fp:
            for line in fp:
                lista_municipios.append(line)
    except:
        pass
    return lista_municipios

# =======================================================================

def obterColunas(arquivo):
    colunas_tabela = []
    colunas_tabela.append('Município')
    fim_linha = 0
    with open(arquivo, 'r', encoding="utf-8") as fp:
        for line in fp:
            posicao_inicial = 0
            while line[posicao_inicial] != "\n":
                posicao_inicial += 1

                posicao_atual = line.find("\tDATA\t", posicao_inicial)
                posicao_aux = posicao_atual

                # Representa a posicao_atual não está vazia (if posicao_atual != -1:)
                if posicao_atual != -1:
                    while line[posicao_aux - 1] != "\t":
                        posicao_aux -= 1
                    if (line[posicao_aux:posicao_atual] not in colunas_tabela):
                        colunas_tabela.append(line[posicao_aux:posicao_atual])

                posicao_inicial = posicao_atual
                posicao_atual = ""
    fp.close()
    return colunas_tabela

# =======================================================================

def abrirArquivo():
    arquivo = filedialog.askopenfilename(initialdir="", title="Escolha um arquivo",
                                     filetypes=(("Arquivo de texto", ".txt"),))
    return arquivo

# =======================================================================

def tabulartransferencias():
    arquivo = abrirArquivo()
    colunasdatabela = obterColunas(arquivo)
    colunasdatabela.append(arquivo)

    return colunasdatabela

# =======================================================================

def organizararquivotransferenciatxt():

    # files = [('All Files', '*.*'), ('Python Files', '*.py'), ('Text Document', '*.txt')]
    files = [('Text Document', '*.txt')]
    arqnome = filedialog.asksaveasfilename(filetypes=files, defaultextension=files)
    arqnome_aux = arqnome

    try:
        os.remove(arqnome)
    except:
        pass

    arqnome = arqnome_aux

    with open("transferencias.txt", "r") as arq:
        linhasDoArquivo = arq.readlines()
        registro = ""
        for linha in linhasDoArquivo:  # linha começa da posição 0
            registro = registro + linha.replace("\n", "\t")
            if linha[0:11] == '[bb.com.br]':
                registro = linha.replace("\n", "", 1)
            if linha[0:13] == 'Transparência':
                open(arqnome, "a", encoding="utf-8").write(registro + "\n")
                registro = linha[13:24] + "\t"
    arq.close()

    try:
        os.remove("transferencias.txt")
    except:
        pass

    messagebox.showinfo("Tranferência BB", "Finish!")


# =======================================================================

def obtertransferenciabb(municipios_selecionados, uf, data_inicial, data_final):

    try:
        options = webdriver.ChromeOptions()
        options.add_argument("headless")
        #navegador = webdriver.Chrome(options=options)
        navegador = webdriver.Chrome()
    except:
        messagebox.showinfo("Tranferência BB", "Atenção: provavelmente seu webdriver.Chrome esteja desatualizado."
                                               "Logo, verifique a versão do seu navegador e baixe o arquivo anteriormente"
                                               "citado, na pasta onde se encontra este APP (onde está instalado o Python.exe)"
                                               "Veja a versão do google digitando:"
                                               "chrome://settings/help. Depois baixe o webdrive correspondente.")

    for municipio in municipios_selecionados:
        navegador.get("https://www42.bb.com.br/portalbb/daf/beneficiario,802,4647,4652,0,1.bbx")
        navegador.find_element(By.XPATH, '// *[ @ id = "formulario:txtBenef"]').send_keys(municipio)
        navegador.find_element(By.XPATH, '// *[ @ id = "formulario:txtBenef"]').send_keys(Keys.ENTER)
        navegador.find_element(By.XPATH, '//*[@id="formulario:comboBeneficiario"]').send_keys(municipio + " - " + uf)
        navegador.find_element(By.XPATH, '//*[@id="formulario:dataInicial"]').send_keys(data_inicial)
        navegador.find_element(By.XPATH, '//*[@id="formulario:dataFinal"]').send_keys(data_final)
        navegador.find_element(By.XPATH, '//*[@id="formulario:dataFinal"]').send_keys(Keys.ENTER)

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')

        transferencias = pyperclip.paste()
        transferencias = transferencias.replace('\n', "")

        with open("transferencias.txt", "a") as arq:
            arq.write(transferencias)
        arq.close()

    navegador.close()

    organizararquivotransferenciatxt()

# =======================================================================