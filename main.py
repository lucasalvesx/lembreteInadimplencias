# ------------------------------------------------------------------
# Autor: Lucas Alves
# Data: 02/2025
# ------------------------------------------------------------------


#importing libs
import pandas as pd
import yagmail 
import time
import os
from dotenv import load_dotenv

#loading environment variables (credentials stored in .env file)
load_dotenv(dotenv_path="./.env")
USER_EMAIL = os.getenv("USER_EMAIL")
USER_PASSWORD = os.getenv("USER_PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
SMTP_SERVER = os.getenv("SMTP_SERVER")
IMAP_PORT = os.getenv("IMAP_PORT")
SMTP_PORT = int(os.getenv("SMTP_PORT"))#converting datatype for int
SMTP_SSL = os.getenv("SMTP_SSL") == "True"


#handling errors in credentials importing
if not USER_EMAIL or not USER_PASSWORD:
    print("Erro ao importar credenciais de email. Fale com o programador do sistema.")
    exit()

#reading excel file with pandas
try:
    df = pd.read_excel("./clientesInadimplencia.xlsx")
except Exception as e:
    print(f"Erro ao ler arquivo Excel: {e}. Se necessário contatar o programador do sistema.")
    exit() ##WORKS UNTIL HERE

#checking if row exists
if "email" not in df.columns:
    print("Coluna 'email' não encontrada no arquivo Excel. Favor verificar o arquivo.")
    exit()
    
#extracting list of emails
ListaDestinatarios = df["email"].tolist()
print(ListaDestinatarios)

#Enviar email com yagmail

#configurando autenticações:
try:
    yag = yagmail.SMTP(user=USER_EMAIL, password=USER_PASSWORD, host=SMTP_SERVER, port=SMTP_PORT, smtp_ssl=SMTP_SSL)
except Exception as e:
    print(f"Erro ao logar no email: {e}")
    exit()
#nao havendo erros, enviar email
try:
    yag.send(
    bcc = ListaDestinatarios,
    subject = 'Lembrete pagamento honorários',
    contents = 'Você está recebendo essa mensagem como teste. O sistema está próximo de ser implementado.' 
    )
    print("Emails enviados com sucesso!")
except Exception as e:
    print(f"Erro ao enviar email: {e}")
    
# setting batch size (amount of emails sent at a time):
batch_size = 100

#sending in batches
for i in range(0, len(ListaDestinatarios), batch_size):
    batch = ListaDestinatarios[i:i + batch_size]
    try:
        yag.send(
            cc = 'financeiro2@micdigital.com.br',
            bcc= batch,
            subject='Lembrete pagamento honorários',
            contents='Lembramos que consta um pagamento em aberto'
        )
        print(f"Emails enviados com sucesso para o lote {i // batch_size + 1}")
        time.sleep(3)  # Pausa de 3 segundos entre os envios às batches
    except Exception as e:
        print(f"Erro ao enviar email para o lote {i // batch_size + 1}: {e}")

print("Envio de emails finalizado com sucesso!")
