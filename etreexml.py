import os
import pandas as pd
import xml.etree.ElementTree as et
from xml.dom import minidom


def ler_arquivos(arquivo, valores):
    with open(f'nfs/{arquivo}', 'rb') as arquivo_xml:
        try:
            print(f'Lendo o arquivo {arquivo}')
            root = et.parse(arquivo_xml).getroot()
            nsNFE = {'ns': "http://www.portalfiscal.inf.br/nfe"}
            if root.find('ns:NFe', nsNFE) is None:
                nfe_info ='ns:infNFe'
            else:
                nfe_info= 'ns:NFe/ns:infNFe'
            
            numero_nfe = root.find(f'{nfe_info}/ns:ide/ns:nNF', nsNFE).text
            emissor_nfe = root.find(f'{nfe_info}/ns:emit/ns:xNome', nsNFE).text
            dest_nfe = root.find(f'{nfe_info}/ns:dest/ns:xNome', nsNFE).text
            transportadora = root.find(f'{nfe_info}/ns:transp/ns:transporta/ns:xNome', nsNFE).text
            if root.find(f'{nfe_info}/ns:transp/ns:vol', nsNFE) is None:
                peso_liquido = 'Não informado'
                peso_bruto = 'Não informado'
                
            else: 
                peso_liquido =  root.find(f'{nfe_info}/ns:transp/ns:vol/ns:pesoL', nsNFE).text
                peso_bruto = root.find(f'{nfe_info}/ns:transp/ns:vol/ns:pesoB', nsNFE).text
            print(f'Nota Fiscal: {numero_nfe}\n Emissor: {emissor_nfe}\n Destinaratio: {dest_nfe}\n Transportadora: {transportadora}\n Peso Liquido: {peso_liquido}\n Peso Bruto: {peso_bruto}')
            print()
            valores.append([numero_nfe, emissor_nfe, dest_nfe, transportadora, peso_liquido, peso_bruto]) 
                 
        except Exception as e:
            print(e)
            print(f'Deu erro no {arquivo}\n')

lista_arquivos = os.listdir('nfs')

colunas = ['Nota Fiscal', 'Emissor', 'Destinatario', 'Transportadora', 'Peso Liquido', 'Peso Bruto']
valores = []

for arquivo in lista_arquivos:
    ler_arquivos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFiscais.xlsx', index=False)