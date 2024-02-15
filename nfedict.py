import xmltodict
import os
import pandas as pd

def ler_arquivos(arquivos, valores):

    with open(f'nfs/{arquivos}', 'rb') as arquivo_xml:
        root = xmltodict.parse(arquivo_xml)
        if 'NFe' in root:
            nfe_infos = root['NFe']['infNFe']
        else:
            nfe_infos = root['nfeProc']['NFe']['infNFe']
        numero_nfe = nfe_infos['ide']['nNF']
        emissor_nfe = nfe_infos['emit']['xNome']
        dest_nfe = nfe_infos['dest']['xNome']
        transportadora = nfe_infos['transp']['transporta']['xNome']
        if 'vol' in nfe_infos['transp']:
            peso_liquido = nfe_infos['transp']['vol']['pesoL']
            peso_bruto = nfe_infos['transp']['vol']['pesoB']
        else:
            peso_liquido = 'Não informado'
            peso_bruto = 'Não informado'

        valores.append([numero_nfe, emissor_nfe, dest_nfe, transportadora, peso_liquido, peso_bruto])

lista_arquivos = os.listdir('nfs')
colunas = ['Nota Fiscal', 'Emissor', 'Destinatario', 'Transportadora', 'Peso Liquido', 'Peso Bruto']
valores = []
for arquivos in lista_arquivos:
    ler_arquivos(arquivos, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFiscaisToDict.xlsx')