import win32com.client
import pandas as pd
import numpy as np
from pathlib import Path
import os

path_file =  os.path.abspath(str(Path.home()) +'\OneDrive - HEALTHINK\Inteligência\Aplicações\Portal Logado\Conta Nova.xlsx')
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = outlook.Accounts


os.system('cls')

print("Contas disponíveis:")
for i, account in enumerate(accounts):
    print(f"{i+1}. {account.DisplayName}")

account_index = int(input("\nEscolha o número da conta que deseja usar: ")) - 1
selected_account = accounts[account_index]
os.system('cls')

df = pd.read_excel(path_file)
def enviar_email_via_outlook(account):
        outlook = win32com.client.Dispatch("Outlook.Application")
        for index, row in df.iterrows():
            destinatario = row['EMAIL']
            senha = row['SENHA']
            mail = outlook.CreateItem(0) 
            mail.Recipients.Add(destinatario)

            msg = f"""
Olá,

Sua conta para acesso ao portal logado foi criada. 
Segue as credenciais de acesso:

Seu login é: { destinatario }
Sua nova senha é: { senha }

Você pode acessar sua conta através do link abaixo:
Acesso: https://www.htksaude.com.br/corporate/login

Este é um e-mail automatico, não responda.

Atenciosamente,
Equipe Healthink
                """


            mail.Subject = '[Portal Logado] - Criação usuário - Healthink'
            mail._oleobj_.Invoke(*(64209,0,8,0,account))
            mail.Body = msg

            mail.Send()
            print(f'{index+1}) [Enviado] - {destinatario}')


enviar_email_via_outlook(selected_account)
    


