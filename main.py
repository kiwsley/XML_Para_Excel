import xmltodict
import os
import json
import pandas as pd

#Criando função para pegas as informações do da nota fiscal

def pegar_infos(nome_arquivo,valores):
    #print(f"pegou as informações {nome_arquivo}")
    #Abrindo o arquivo
    with open(f'nfs/{nome_arquivo}',"rb") as arquivo_xml:
        #Tranformando em dicionario
        dic_arquivo = xmltodict.parse(arquivo_xml)
       # print(json.dumps(dic_arquivo, indent=4))

    #Buscando as informações para os diferente tipos de tags
    # Os blocos try e except, foram criados com o objetivos de comparar os arquivos caso tenham divergências nas tags   
    try:
        if "NFe" in dic_arquivo:    
            infos_nf = dic_arquivo["NFe"]['infNFe']
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]['infNFe']
        numero_nota= infos_nf["@Id"]
        empresa_emissora=infos_nf['emit']['xNome']
        nome_cliente=infos_nf['dest']['xNome']
        endereco=infos_nf['dest']['enderDest']
        if 'vol' in infos_nf['transp']:
             peso=infos_nf['transp']['vol']['pesoB']
        else:  
            peso="não informado"
        #print(numero_nota, empresa_emissora,nome_cliente,endereco,peso, sep="\n")
        #

        #Inserindo as informações no dicionario
        valores.append([numero_nota, empresa_emissora,nome_cliente,endereco,peso])

    except Exception as e:
        print(e)
        print (json.dumps(dic_arquivo, indent=4))
            

        
#Listando os arquivos
lista_arquivos = os.listdir("nfs")
##Criando as colunas e valores do arquivo
colunas = ['numero_nota',"empresa_emissora","nome_cliente","endereço","peso"]
valores=[]

#passando as informações para a função
for arquivo in lista_arquivos:
    pegar_infos(arquivo,valores)

#Criando o arquivo excel    
tabela= pd.DataFrame(columns=colunas,data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)
