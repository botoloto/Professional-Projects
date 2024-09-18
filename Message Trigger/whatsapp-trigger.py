import pandas as pd
from pathlib import Path
import os
import requests
import time
import json

erros_executivos, enviados_executivos,enviados_cliente,erros_cliente = 0,0,0,0

# Configuração de endpoint e headers para conexão com o disparador de mensagem
endpoint = "https://exemplo.tunnel.zaplink.net/mensagem/envio"
headers = {
    "Authorization": "Bearer Token1", 
    "Content-Type": "application/json"
}

# Referenciando o arquivo que contém as informações de contato
file_disparo = os.path.join(Path.home(), 'Downloads\Acao.xlsx')
df = pd.read_excel(file_disparo)

for index, row in df.iterrows():
    #Pausa para evitar banimento de spam de msg
    if index != 0:  
        time.sleep(10) 

    # Criando variaveis com o numero e a mensagem única que sera enviada para cada pessoa individualmente
    tipo = row['Tipo de Acao']
    responsavel = row['Responsavel']
    nome_executivo = row['Executivo']
    cliente = row['Cliente']
    nome_cliente = row['Responsavel Cliente']

    msg_executivo = f"""Olá, {nome_executivo}.

A ação de {tipo} em parceria com a {cliente} foi finalizada. 
Gostaria de reforçar que sua opinião como executivo é importante para compararmos com a avaliação do cliente afim de proporcionar experiencias ainda melhores.

Então gostaríamos do seu feedback! 
Contamos com a sua opinião: https://app.pipefy.com/public/form/8PyZHbW_

Obrigado pela participação, até a próxima! 😉
{responsavel}"""

    msg_cliente = f"""Olá, {nome_cliente}. 

A a ação de {tipo}, uma realização da Healthink em parceria com a {cliente} foi realizada recentemente. 
Gostariamos de reforçar que sua opinião é muito valiosa para nós, pois nos ajuda a proporcionar experiencias ainda melhores.

Se possível, gostaríamos do seu feedback! 
É super rápido e faz toda a diferença para nós. Por favor, clique no link e nos conte a sua opinião: https://app.pipefy.com/public/form/sMjtuhKr

Continuamos com o compromisso de conectar vidas e cuidar de futuros.
Agradecemos muito por compartilhar suas impressões conosco.

Até a próxima! 😉
{responsavel}"""

    nmr_executivo = str(row['Telefone Executivo'])
    nmr_cliente = str(row['TelefoneEmpresa'])
    
    #Criação do corpo responsável pelo envio
    body_executivo = json.dumps({ 
        "number": nmr_executivo,
        "body": msg_executivo
    })

    body_cliente = json.dumps({ 
        "number": nmr_cliente,
        "body": msg_cliente
    })
    
    #Envio da mensagem com o retorno do status do envio
    resultado_executivo = requests.post(endpoint, headers=headers, data=body_executivo)
    print(f'\n\nEnvio executivos: {resultado_executivo}')
    time.sleep(10) 
    resultado_cliente = requests.post(endpoint, headers=headers, data=body_cliente)
    print(f'\n\nEnvio cliente: {resultado_cliente}')

