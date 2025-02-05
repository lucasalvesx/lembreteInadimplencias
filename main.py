import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# smtp config
smtp_server = "scca2.icloudzinc.com"
smtp_port = 465
email_remetente = "comunicado@micdigital.com.br"
senha = "####"

# linking excel sheet
clientes_df = pd.read_excel("clientesInadimplencia.xls")

# function for sending messages
def enviar_email(cliente_email, cliente_nome):
    assunto = "Lembrete pagto honorários"
    corpo = f"""
    Olá, cliente {cliente_nome},
     Gostaríamos de lembrá-lo(a) que o pagamento dos honorários devidos ainda não foi efetuado.
    Pedimos que regularize sua pendência o mais breve possível para evitar a interrupção dos serviços.

    Atenciosamente,
    MIC Contabilidade Digital
    """
    
    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = cliente_email
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))
    
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_remetente, senha)
            server.sendmail(email_remetente, cliente_email, msg.as_sring())
            print(f"Enviado para {cliente_nome} ({cliente_email})")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {cliente_nome}: {e}")
        
# iterar na lista de clientes e executar funçao de envio
for _, cliente in clientes_df.iterrows():
    enviar_email(cliente['email'], cliente['nome'])