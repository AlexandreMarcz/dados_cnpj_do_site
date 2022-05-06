# Scrypt para obter dados de CNPJ salvo em planilha.
# O nome da coluna com os dados da planilha deve ser 'CNPJ da Carga'
# Uso da api do site www.receitaws.com.br


import json
import requests
import pandas
import time
import random

from datetime import datetime as dt



arquivo_leitura = pandas.read_excel('planilha_com_os_cnpjs.xlsx', sheet_name='Planilha1')

# Para determinar o numero de interacoes na planilha

tam_total = len(arquivo_leitura.index)

# O laco for a seguir vai ler o cnpj na planilha, buscar os dados
# pela API e salvar na planilha

for x in range(0,tam_total): 
    
    
    try:
	    y = str(int(arquivo_leitura.loc[x,'CNPJ da Carga']))
	
    except ValueError:
	    print('Carga sem CNPJ')
    
    else:
    
        print(y)
    
        if len(y)== 10:
            url = "https://www.receitaws.com.br/v1/cnpj/"+'0000'+y
    
        if len(y)== 11:
            url = "https://www.receitaws.com.br/v1/cnpj/"+'000'+y
   
        if len(y)== 12:
            url = "https://www.receitaws.com.br/v1/cnpj/"+'00'+y
    
        if len(y)== 13:
            url = "https://www.receitaws.com.br/v1/cnpj/"+'0'+y
    
        else:
            url = "https://www.receitaws.com.br/v1/cnpj/"+y
        
    
        print(url)
    
        response = requests.get(url)
        data = json.loads(response.content)
        nome_empresa = data.get("nome")
        logradouro = data.get('logradouro')
        numero = data.get('numero')
        cep = data.get('cep')
        bairro = data.get('bairro')
        municipio = data.get('municipio')
        email = data.get('email')
        telefone = data.get('telefone')
    
    
        arquivo_leitura.loc[x,'razao_social'] = nome_empresa
        arquivo_leitura.loc[x,'logradouro'] = logradouro
        arquivo_leitura.loc[x,'numero'] = numero
        arquivo_leitura.loc[x,'cep'] = cep
        arquivo_leitura.loc[x,'bairro'] = bairro
        arquivo_leitura.loc[x,'municipio'] = municipio
        arquivo_leitura.loc[x,'email'] = email
        arquivo_leitura.loc[x,'telefone'] = telefone
        arquivo_leitura.to_excel('cadastroos - se.xlsx',sheet_name = 'Planilha1', index=False, engine='openpyxl')
    
    
    
        print(arquivo_leitura.loc[x,'telefone'])
        print('passo ',x ,' de ',tam_total )
    
    # Como a API gratuiraso permite 3 acessos por minuto, foi colocado um 
    # temporizador para ela consultar dentro desse tempo

        tempo_aleatorio = random.randint(20,21)
        time.sleep(tempo_aleatorio)
    
print('Terminou, trabalho executado')
   
    

    

    


