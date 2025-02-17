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
import tkinter as tk #native module for GUI
from tkinter import filedialog #module for file dialog

#creating GUI window
root = tk.Tk()
root.title("Enviar mensagens")
root.geometry("300x200")
#runnng tkinter main loop
root.mainloop()

# Add instructions label
label = tk.Label(root, text="Por favor, selecione o arquivo Excel com os dados.", padx=10, pady=10)
label.pack()

#function for file dialog(insert files by browsing on user local pc)
def browse_file():
    """Faça o upload do arquivo desejado."""
    file_path = filedialog.askopenfilename (
        title="Selecione o arquivo Excel",
        filetypes=(("Excel files", "*.xlsx"))
    )
    print(f"Arquvio selecionado {file_path}")
    return file_path
    
# Add a button to insert
select_button = tk.Button(root, text="Selecionar Arquivo Excel", command={browse_file})
select_button.pack(padx=10, pady=10)

#function for files processing
def process_file(file_path):
    """Processa o arquivo Excel."""
    try:
        df_lembrete = pd.read_excel(file_path, sheet_name="Lembrete")
        df_cobranca = pd.read_excel(file_path, sheet_name="Cobranca")
    except Exception as e:
        print(f"Erro ao ler arquivo Excel: {e}. Se necessário contatar o programador do sistema.")
    return df_lembrete, df_cobranca

#loading environment variables (credentials stored in .env file)
load_dotenv(dotenv_path="./app/.env")
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
    df_lembrete = pd.read_excel("./app/clientesInadimplencia.xlsx", sheet_name="Lembrete")
    df_cobranca = pd.read_excel("./app/clientesInadimplencia.xlsx", sheet_name="Cobranca")
except Exception as e:
    print(f"Erro ao ler arquivo Excel: {e}. Se necessário contatar o programador do sistema.")
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
    for i in range(0, len(listaDestinatarios), batch_size): #iterates over read list of emails
        batch = listaDestinatarios[i:i + batch_size] #slices list and sets batch value 
        
        try: 
            yag.send(
                cc='USER_EMAIL',
                bcc=batch,
                subject=subject,
                contents=content
            )
            print(f"Emails enviados com sucesso para o lote {i // batch_size + 1}")
            time.sleep(3)  # waits 3 seconds to start sending for the next batch
        except Exception as e:
            print(f"Erro ao enviar email para o lote {i // batch_size + 1}: {e}")

#sending reminder email
print("Enviando lembretes...")
enviar_emails(envio_lembretes, "Lembrete de vencimento", "TESTE SISTEMA - Cliente, lembramos do vencimento")

#sending expired payment email 
print("Enviando cobranças...")
enviar_emails(envio_cobranca, "Pagamento vencido", "TESTE SISTEMA - Cliente, seu pagamento consta em atraso.")

## end function

print("Envio de emails finalizado com sucesso!")
