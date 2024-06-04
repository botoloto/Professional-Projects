import pandas as pd
from pathlib import Path
import os
import requests
import time
import json


erros, enviados = 0

# Configuração de endpoint e headers para conexão com o disparador de mensagem
endpoint = "https://0000000000000.tunnel.zaplink.net/mensagem/envio"
headers = {
    "Authorization": "Bearer Token1", 
    "Content-Type": "application/json"
}

# Referenciando o arquivo que contém as informações de contato
file_disparo = os.path.join(Path.home(), 'Downloads\Mensagens não recebidas - Abril-2024.xlsx')
df = pd.read_excel(file_disparo)

for index, row in df.iterrows():
    #Pausa para evitar banimento de spam de msg
    if index != 0:  
        time.sleep(10) 

    # Criando variaveis com o numero e a mensagem única que sera enviada para cada pessoa individualmente
    mensagem_personalizada = f"Olá, {row['Nome Pessoa']}\nTemos o prazer de informar que seu atendimento foi concluído a alguns dias.\nAgradecemos por escolher nossos serviços.\n\nPara nos ajudar a aprimorar nossa qualidade, gostaríamos de solicitar sua opinião através de uma breve pesquisa de satisfação.\nSuas respostas são valiosas para nós.\n\nPor favor, acesse a pesquisa pelo seguinte link,  e se o link abaixo não estiver ativo, por favor, copie e cole o seguinte endereço em seu navegador web para acessar a pesquisa, ou responda com qualquer texto para ativar o link clicável:\n{row['URL']}\n\nSeu feedback é fundamental para continuarmos melhorando.\n\nAtenciosamente,\nMaria Victoria S. L. Batista"
    numero = row['Telefone']
    
    #Criação do corpo responsável pelo envio
    body = json.dumps({ 
        "number": numero,
        "body": mensagem_personalizada
    })

    #Envio da mensagem com o retorno do status do envio
    resultado = response = requests.post(endpoint, headers=headers, data=body)

    #Armazenamento do status do envio
    if 'OK' in resultado:
        print(f"ENVIADO - {row['Nome Pessoa']}: {resultado} - {row['Telefone']}")
        enviados += 1
        df.at[index, 'STATUS'] = 'ENVIO'
    else:
        print(f"ERRO - {row['Nome Pessoa']}: {resultado} - {row['Telefone']}")
        erros += 1
        df.at[index, 'STATUS'] = 'ERRO'
    
#Etapa na qual salva o status de todos os envios para os números na mesma planilha anexada
print(f'\n\nENVIOS: {enviados}\nERROS: {erros}')
df.to_excel('Status Envios.xlsx',index=False)
