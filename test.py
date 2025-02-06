import pandas as pd
import yagmail 

df = pd.read_excel(".\clientesInadimplencia.xlsx")
ListaDestinatarios = df

print(df.head())

yag = yagmail.SMTP('comunicado@micdigital.com.br', '####')
yag.send(
    to = ListaDestinatarios
    subject = 'Lembrete pagamento honorários'
    contents = 'Lembramos que consta um pagamento em aberto' 
    )

# error handle

