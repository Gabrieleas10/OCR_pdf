# -*- coding: utf-8 -*-

# IMPORTAÇÃO BIBLIOTECAS
print("Importando bibliotecas...\n")
import PyPDF2
import pandas as pd
import getpass
import os
from geopy.distance import geodesic
from tqdm import tqdm
import numpy as np
import xlrd
import utm
print('Finalizado!\n')


# IDENTIFICAÇÃO DO USUARIO DO PC
print("Identificando o Usuário...\n")
br = getpass.getuser()
print('Finalizado!\n')

# IDENTIFICAÇÃO DOS PATHS DOS DADOS
print("Identificando o caminho dos arquivos...\n")
base_path = 'C:\\Users\\' + br + '\\Enel Spa\\NIP Ceará - forcoe_backup\\PLANEJAMENTO_AT_MT\\Cartas\\2020\\Mala Direta\\PA.xlsx'
sheet_pa = 'Parecer de Acesso'

pa_path = 'C:\\Users\\' + br + '\\Enel Spa\\NIP Ceará - forcoe_backup\\PLANEJAMENTO_AT_MT\\Cartas\\2020\\Mala Direta\\SA'

barras_path = 'C:\\Users\\' + br + '\\Enel Spa\\NIP Ceará - forcoe_backup\\PLANEJAMENTO_AT_MT\\Estudos_MT\\Ciclo_2019\\Quadrículas\\Curto_CE_Oficial.xlsx'
print('Finalizado!\n')

# IMPORTAÇÃO DAS BASES DE DADOS
print("Carregando os dados...\n")
base = pd.read_excel( base_path , sheet_pa)

barras = pd.read_excel(barras_path)

lista_pa = os.listdir(pa_path)
print('Finalizado!\n')

# FUNÇÕES AUXILIARES
print("Construindo funções auxiliares...\n")
tqdm.pandas()

def distance(lat , lon):
    try:
        lat = str(lat)
        lon = str(lon)
        if lat.find('-')!= -1:
            latpa = lat.replace('°','')
            lonpa = lon.replace('°','')
            latpa = latpa.replace(',','.')
            lonpa = lonpa.replace(',','.')
            dist = []
        elif lat.find('mS'):
            latpa = lat.replace('mS','')
            lonpa = lon.replace('mE','')
            latpa = latpa.replace(',','.')
            lonpa = lonpa.replace(',','.')
            coord_utm = utm.to_latlon(lonpa , latpa, 24, 'M')
            latpa = coord_utm[0]
            lonpa = coord_utm[1]
            dist = []
        for i in tqdm(range(0 , len(barras['X1']),1)):
            coord_1 = (float(latpa) ,float(lonpa))
            coord_2 = (barras['X1'][i] , barras['Y1'][i] )
            km = geodesic(coord_1 ,coord_2 ).km
            dist.append(km)
        index_min = np.argmin(dist)
        kmd = dist[index_min]
        sed = barras['Subestaçăo'][index_min]
        feeder = barras['Alimentador'][index_min]
        cd = barras['Código da barra'][index_min]
    except:
        sed = ''
        feeder = ''
        cd = '' 
        kmd = ''
    return sed,feeder,cd,kmd
print('Finalizado!\n')
    
# OCR DOS DOCUMENTOS DE ACESSO EM .PDF E TRANSFORMAÇÃO PARA STRING
print("Lendo os arquivos pdf...\n")   
for i in range(0 , len(lista_pa),1):
    pdfFileObj = open( pa_path +'\\'+ lista_pa[i] , 'rb')

    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)  

    pageObj = pdfReader.getPage(0) 

    texto = pageObj.extractText()

    pdfFileObj.close()
    
    if texto != "":
# TRATAMENTO DA STRING
        texto = texto.replace('\n','')
        texto = texto.replace(' ','')
    
        titular = texto[texto.find('TitulardaUC:')+12:texto.find('Rua/Av.:')]
    
        tensao = texto[texto.find('Tensãodeatendimento(V)')+22:texto.find('Tipodeconexão')]
    
        lat = texto[texto.find('Latitude:')+9:texto.find('Longitude:')]
    
        lon = texto[texto.find('Longitude:')+10:texto.find('Potênciainstalada')]
        
        rede_mt = distance(lat , lon)
        
        sed = rede_mt[0]
        feeder = rede_mt[1]
        cd = rede_mt[2]
        ramal = rede_mt[3]
    
        potencia = texto[texto.find('Potênciainstaladadegeração(kW):')+31:texto.find('TipodaFontedeGeração')]
    
        cidade = texto[texto.find('Cidade:')+7:texto.find('E-mail:')]
    
        uc = texto[texto.find('CódigodaUC:')+11:texto.find('Grupo')]
    
        parte2 = texto[texto.find('SolicitanteNome/ProcuradorLegal:'):texto.find('AssinaturadoResponsável')]
    
        rep_legal = parte2[parte2.find('SolicitanteNome/ProcuradorLegal:')+32:parte2.find('Telefone:')]
    
        telefone = parte2[parte2.find('Telefone:')+9:parte2.find('E-mail:')]
    
        os = lista_pa[i][0:lista_pa[i].find('_')]
    
        data = lista_pa[i][lista_pa[i].find('_')+1:lista_pa[i].find('.pdf')]
    
        num_points = parte2[parte2.find('@'):].count('.')
        loc_points = parte2.find('.')
    
        if num_points == 0:
            email = 0
        elif num_points == 1:
            email = parte2[parte2.find('E-mail:')+7:parte2.find('.')+4]
        elif num_points == 2:
            email = parte2[parte2.find('E-mail:')+7:parte2.find('.' , loc_points+1)+3]
      
        nova_solic = [rep_legal , titular , 'Geração' , potencia , uc , potencia , lat , lon , tensao ,
                      sed , cidade , feeder , cd , ramal , data , os , 'Digital' , rep_legal , telefone, email ]
        
        list_colunas = list(base.columns)
        print('Finalizado!\n')

# INSERINDO NA BASE DE DADOS
        print("Inserindo os dados na planilha de controle...\n")
        row = pd.Series( nova_solic , list(base.columns))

        base = base.append([row],ignore_index=True)
    elif texto == "":
        os = lista_pa[i][0:lista_pa[i].find('_')]
    
        data = lista_pa[i][lista_pa[i].find('_')+1:lista_pa[i].find('.pdf')]
        
        nova_solic = ["" , "" , 'Geração' , "" , "" , "" , "" , "" , "" ,
                      "" , "" , "" , "" , "" , data , os , 'Digital' , "" , "", "" ]
        
        list_colunas = list(base.columns)
        
        print("Inserindo os dados na planilha de controle...\n")
        row = pd.Series( nova_solic , list(base.columns))

        base = base.append([row],ignore_index=True)
        
print("Salvando a planilha...\n")
# SALVANDO O ARQUIVO NO SHAREPOINT
final_path = 'C:\\Users\\' + br + '\\Enel Spa\\NIP Ceará - forcoe_backup\\PLANEJAMENTO_AT_MT\\Cartas\\2020\\Mala Direta\\PA.xlsx' 
base.to_excel(final_path , header=True, index=False , sheet_name= sheet_pa , engine='xlsxwriter')
print("Finalizado...\n")

print("Excluindo as Solicitações de Acesso...\n")
for file in os.scandir(pa_path):
    if file.name.endswith(".pdf"):
        os.unlink(file.path)
print("Finalizado...\n")
print("Concluído com Sucesso!!!\n")