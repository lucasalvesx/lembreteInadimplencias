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

# Global variables for storing dataframes
df_lembrete = None
df_cobranca = None

# Loading environment variables (credentials stored in .env file)
load_dotenv(dotenv_path="./app/.env")
USER_EMAIL = os.getenv("USER_EMAIL")
USER_PASSWORD = os.getenv("USER_PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
SMTP_SERVER = os.getenv("SMTP_SERVER")
IMAP_PORT = os.getenv("IMAP_PORT")
SMTP_PORT = int(os.getenv("SMTP_PORT"))#converting datatype for int
SMTP_SSL = os.getenv("SMTP_SSL") == "True"

# Handling errors in credentials importing
if not USER_EMAIL or not USER_PASSWORD:
    print("Erro ao importar credenciais de email. Fale com o programador do sistema.")
    exit()
    
# Function to process selected file
def process_file(file_path):
    global df_lembrete, df_cobranca
    try:
        df_lembrete = pd.read_excel(file_path, sheet_name="Lembrete")
        df_cobranca = pd.read_excel(file_path, sheet_name="Cobranca")
        print("Arquivo lido com sucesso!")
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")    
# end function process_file
          
# Function to open file dialog and insert files by browsing on user local pc
def browse_file():
    """Faça o upload do arquivo desejado."""
    global enviar_button
    file_path = filedialog.askopenfilename (
        title="Selecione o arquivo Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        label_status.config(text=f"Arquivo selecionado: {file_path}")
        process_file(file_path)
        enviar_button.config(state="normal")
# end function browse_file

# function to send emails  
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
                cc=USER_EMAIL,
                bcc=batch,
                subject=subject,
                contents=content
            )
            print(f"Emails enviados com sucesso para o lote {i // batch_size + 1}")
            time.sleep(3)  # waits 3 seconds to start sending for the next batch
        except Exception as e:
            print(f"Erro ao enviar email para o lote {i // batch_size + 1}: {e}")
## end function enviar_emails
        
# Function to send messages
def enviar_mensagens():
    global df_lembrete, df_cobranca  

    if df_lembrete is None or df_cobranca is None:
        print("Erro: Nenhum arquivo foi carregado.")
        label_status.config(text="Erro: Nenhum arquivo carregado.", fg="red")
        return

    try:
        envio_lembretes = df_lembrete["email"].dropna().tolist()
        envio_cobranca = df_cobranca["email"].dropna().tolist()

        print("Enviando lembretes...")
        enviar_emails(envio_lembretes, "Lembrete de vencimento", "TESTE SISTEMA - Cliente, lembramos do vencimento")

        print("Enviando cobranças...")
        enviar_emails(envio_cobranca, "Pagamento vencido", "TESTE SISTEMA - Cliente, seu pagamento consta em atraso.")

        print("Envio de emails finalizado com sucesso!")
        label_status.config(text="Mensagens enviadas com sucesso!", fg="green")

    except Exception as e:
        print(f"Erro ao enviar mensagens: {e}")
        label_status.config(text="Erro no envio de mensagens!", fg="red")
# end function enviar_mensagens

# end of main.py (LOGICAL PART)

# configuring tkinter GUI (widgets)
root = tk.Tk()
root.title("Enviar mensagens")
root.geometry("400x250")

# label for instructions
label = tk.Label(root, text="Por favor, selecione o arquivo Excel com os dados.", padx=10, pady=10)
label.pack()

# button to select file 
select_button = tk.Button(root, text="Selecionar Arquivo Excel", command=browse_file)
select_button.pack(padx=10, pady=10)

# status for label
label_status = tk.Label(root, text="", fg="blue")
label_status.pack()

# confirmation button (unsabled until a file is loaded)
enviar_button = tk.Button(root, text="Confirmar Envio", command=enviar_mensagens, state="disabled")
enviar_button.pack(padx=10, pady=10)

# button for quit
quit_button = tk.Button(root, text="Sair", command=root.quit)
quit_button.pack(pady=20)

root.mainloop()

# end of GUI configuration

    




