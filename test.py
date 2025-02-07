#importing libs
import pandas as pd
import yagmail 
import time
import os
from dotenv import load_dotenv

#loading environment variables (credentials stored in .env file)
load_dotenv(dotenv_path="./clientesInadimplencia.xlsx")
USER_EMAIL = os.getenv("USER_EMAIL")
USER_PASSWORD = os.getenv("USER_PASSWORD")

#handling errors in credentials importing
if not USER_EMAIL or not USER_PASSWORD:
    print("Erro ao importar credenciais de email. Fale com o programador do sistema.")
    exit()

#reading excel file with pandas
try:
    df = pd.read_excel(".\clientesInadimplencia.xlsx")
except Exception as e:
    print(f"Erro ao ler arquivo Excel: {e}. Se necessário contatar o programador do sistema.")
    exit()

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
    yag = yagmail.SMTP(USER_EMAIL, USER_PASSWORD)
except Exception as e:
    print(f"Erro ao logar no email: {e}")
    exit()
#nao havendo erros, enviar email
try:
    yag.send(
    to = ListaDestinatarios,
    subject = 'Lembrete pagamento honorários',
    contents = 'Lembramos que consta um pagamento em aberto' 
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
            to=batch,
            subject='Lembrete pagamento honorários',
            contents='Lembramos que consta um pagamento em aberto'
        )
        print(f"Emails enviados com sucesso para o lote {i // batch_size + 1}")
        time.sleep(3)  # Pausa de 3 segundos entre os envios dos batches
    except Exception as e:
        print(f"Erro ao enviar email para o lote {i // batch_size + 1}: {e}")

print("Envio de emails finalizado com sucesso!")
