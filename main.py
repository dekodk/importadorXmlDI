import os
from xml.dom import minidom
import xml.etree.ElementTree as ET
import customtkinter
from tkinter import filedialog
import pymysql.cursors
import tkinter.messagebox as messagebox
import subprocess
import sys

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

#Função para procurar o arquivo xlsx do Excel
def procurar_arquivo2():
    caminho_arquivo2 = filedialog.askopenfilename(title="Selecione um arquivo")
    if caminho_arquivo2:
        end_excel.delete(0, 'end')  # Limpa o conteúdo atual do CTkEntry
        end_excel.insert(0, caminho_arquivo2)  # Insere o novo caminho no CTkEntry

#Função para procurar o arquivo XML da DI
def procurar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(title="Selecione um arquivo")
    if caminho_arquivo:
        end_xml.delete(0, 'end')  # Limpa o conteúdo atual do CTkEntry
        end_xml.insert(0, caminho_arquivo)  # Insere o novo caminho no CTkEntry

def cam_xml():
    caminho_xml = end_xml.get()

    with open(caminho_xml, 'r', encoding='utf-8') as f:
        xml = minidom.parse(f)
        numero_DI = xml.getElementsByTagName("numeroDI")
        data_Registro = xml.getElementsByTagName("dataRegistro")
        carga_PesoLiquido = xml.getElementsByTagName("cargaPesoLiquido")
        local_EmbarqueTotalDolares = xml.getElementsByTagName("localEmbarqueTotalDolares")
        local_EmbarqueTotalReais = xml.getElementsByTagName("localEmbarqueTotalReais")
        frete_TotalMoeda = xml.getElementsByTagName("freteTotalMoeda")
        frete_MoedaNegociadaNome = xml.getElementsByTagName("freteMoedaNegociadaNome")
        seguro_TotalMoedaNegociada = xml.getElementsByTagName("seguroTotalMoedaNegociada")
        seguro_MoedaNegociadaNome = xml.getElementsByTagName("seguroMoedaNegociadaNome")
        denominacao_Acrescimo = xml.getElementsByTagName("denominacao")
        valorReais_Acrescimo = xml.getElementsByTagName("valorReais")
        condicao_VendaIncoterm = xml.getElementsByTagName("condicaoVendaIncoterm")
        condicao_VendaMoedaNome = xml.getElementsByTagName("condicaoVendaMoedaNome")
        numero_Adicao = xml.getElementsByTagName("numeroAdicao")
        numero_SequencialItem = xml.getElementsByTagName("numeroSequencialItem")
        quantidadeDI = xml.getElementsByTagName("quantidade")
        unidadeMedidaDI = xml.getElementsByTagName("unidadeMedida")
        valorUnitarioDI = xml.getElementsByTagName("valorUnitario")

    print(f'Número da DI: {numero_DI[0].firstChild.data}\n')

    print(f'Data de Registro: {data_Registro[0].firstChild.data}\n')

    print(f'Carga Peso Líquido: {carga_PesoLiquido[0].firstChild.data}\n')

    print(f'Valor Merc. Local Embarque total em dólares: {local_EmbarqueTotalDolares[0].firstChild.data}\n')

    print(f'Valor Merc. Local Embarque total em reais: {local_EmbarqueTotalReais[0].firstChild.data}\n')

    print(f'Valor frete moeda local: {frete_TotalMoeda[0].firstChild.data}\n')

    print(f'Moeda do frete: {frete_MoedaNegociadaNome[0].firstChild.data}\n')

    print(f'Seguro: {seguro_TotalMoedaNegociada[0].firstChild.data}\n')

    print(f'Moeda do Seguro: {seguro_MoedaNegociadaNome[0].firstChild.data}\n')

    print(f'Acréscimos: {denominacao_Acrescimo[0].firstChild.data} Valor:{valorReais_Acrescimo[0].firstChild.data}')
    print(f'Acréscimos: {denominacao_Acrescimo[1].firstChild.data} Valor:{valorReais_Acrescimo[1].firstChild.data}')
    print(f'Acréscimos: {denominacao_Acrescimo[2].firstChild.data} Valor:{valorReais_Acrescimo[2].firstChild.data}\n')

    print(f'Incoterm: {condicao_VendaIncoterm[0].firstChild.data}')
    print(f'Moeda VUCV: {condicao_VendaMoedaNome[0].firstChild.data}\n')

    print(f'Adição:{numero_Adicao[0].firstChild.data} Item:{numero_SequencialItem[0].firstChild.data} Quantidade:{quantidadeDI[0].firstChild.data} UNIDADE MED.:{unidadeMedidaDI[0].firstChild.data} VUCV ITEM:{valorUnitarioDI[0].firstChild.data}')
    print(f'Adição:{numero_Adicao[0].firstChild.data} Item:{numero_SequencialItem[1].firstChild.data} Quantidade:{quantidadeDI[1].firstChild.data} UNIDADE MED.:{unidadeMedidaDI[1].firstChild.data} VUCV ITEM:{valorUnitarioDI[1].firstChild.data}')
    print(f'Adição:{numero_Adicao[0].firstChild.data} Item:{numero_SequencialItem[2].firstChild.data} Quantidade:{quantidadeDI[2].firstChild.data} UNIDADE MED.:{unidadeMedidaDI[2].firstChild.data} VUCV ITEM:{valorUnitarioDI[2].firstChild.data}')
    print(f'Adição:{numero_Adicao[1].firstChild.data} Item:{numero_SequencialItem[3].firstChild.data} Quantidade:{quantidadeDI[3].firstChild.data} UNIDADE MED.:{unidadeMedidaDI[3].firstChild.data} VUCV ITEM:{valorUnitarioDI[3].firstChild.data}\n')

#criar telas
janela = customtkinter.CTk()
janela.title("Login")
janela.geometry("700x400")

texto1 = customtkinter.CTkLabel(janela, text="Sistema de exportação de dados XML DI", font=("", 22))
texto1.pack(padx=10, pady=10)

texto2 = customtkinter.CTkLabel(janela, text="Olá Bem vindo!", font=("",19))
texto2.pack(padx=10, pady=10)

end_xml = customtkinter.CTkEntry(janela, placeholder_text="Digite o caminho do arquivo XML", width=330)
end_xml.pack(padx=5, pady=10)
botao2 = customtkinter.CTkButton(janela, text="Procurar", command=procurar_arquivo)
botao2.pack(padx=5, pady=10)


end_excel = customtkinter.CTkEntry(janela, placeholder_text="Digite o caminho para salvar o arquivo Excel", width=330)
end_excel.pack(padx=5, pady=10)
botao3 = customtkinter.CTkButton(janela, text="Procurar", command=procurar_arquivo2)
botao3.pack(padx=5, pady=10)

botao = customtkinter.CTkButton(janela, text="Exportar", command=cam_xml)
botao.pack(padx=20, pady=20)

janela.mainloop()