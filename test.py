#importing libs
import pandas as pd
import yagmail 
import os
from dotenv import load_dotenv

#loading environment variables (credentials stored in .env file)
load_dotenv()
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
    print(f"Erro ao ler arquivo Excel: {e}. S enecessário contatao programador do sistema.")
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



