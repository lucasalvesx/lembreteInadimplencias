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
    df_lembrete = pd.read_excel("./clientesInadimplencia.xlsx", sheet_name="Lembrete")
    df_cobranca = pd.read_excel("./clientesInadimplencia.xlsx", sheet_name="Cobranca")
except Exception as e:
    print(f"Erro ao ler arquivo Excel: {e}. Se necessário contatar o programador do sistema.")
    exit() 

#checking if email row exists in both sheets
if "email" not in df_lembrete.columns or "email" not in df_cobranca.columns:
    print("Coluna 'email' não encontrada no arquivo Excel. Favor verificar o arquivo.")
    exit()
    
#extracting list of emails
envio_lembretes = df_lembrete["email"].dropna().tolist()
envio_cobranca = df_cobranca["email"].dropna().tolist()

#declaring BUT NOT CALLING function to send emails (used further)   
def enviar_emails(listaDestinatarios, subject, content):
    #checking if email row exists in both sheets
    if not listaDestinatarios:
        print("Nenhum endereço encontrado")
        return
    
    #configurando autenticaçoes (login do email destinatario)
    try:
        yag = yagmail.SMTP(user=USER_EMAIL, password=USER_PASSWORD, host=SMTP_SERVER, port=SMTP_PORT, smtp_ssl=SMTP_SSL)
    except Exception as e:
        print(f"Erro ao logar no email: {e}")
        return
        
    batch_size = 100 # setting batch size (amount of emails sent at a time to avoid crashings):

#sending emails in batches
    for i in range(0, len(listaDestinatarios), batch_size):
        batch = listaDestinatarios[i:i + batch_size]
    try:
        yag.send(
            cc = 'financeiro2@micdigital.com.br',
            bcc= batch,
            subject='Teste de implementação MIC',
            contents = 'Você está recebendo essa mensagem como teste. O sistema está próximo de ser implementado.'
        )
        print(f"Emails enviados com sucesso para o lote {i // batch_size + 1}")
        time.sleep(3)  # Pausa de 3 segundos entre os envios às batches
    except Exception as e:
        print(f"Erro ao enviar email para o lote {i // batch_size + 1}: {e}")

#sending reminder email
print("Enviando lembretes...")
enviar_emails(envio_lembretes, "Lembrete de vencimento", "Cliente, lembramos do vencimento")

#sending expired payment email 
print("Enviando cobranças...")
enviar_emails(envio_cobranca, "Pagamento vencido", "Cliente, seu pagamento consta em atraso.")

## end function

print("Envio de emails finalizado com sucesso!")
