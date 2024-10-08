# Envia e-mail aos aniversariantes
def enviaemail(df_aniver):

    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    
    # Conecta ao servidor SMTP    
    smtp_server = 'mail.oneguerbet-group.com'
    smtp_user = 'aniversariantes@guerbet.com'
    server = smtplib.SMTP(smtp_server, 25)
      
    for i, linha in df_aniver.iterrows():
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'Feliz Aniversário!'
        msg['From'] = smtp_user
        msg['To'] = linha['e-Mail']

        # Carrega arquivo HTML
        with open('MensagemAniver.html', 'r') as arquivo:
            html = f"\"" + arquivo.read() + ""
               
        # Anexa o HTML à mensagem
        msg.attach(MIMEText(html, 'html'))
            
        # Envia o e-mail
        server.sendmail(smtp_user, linha['e-Mail'], msg.as_string())
    
    # Fecha a conexão com o servidor SMTP
    server.quit()


# Inicia programa principal
def main():

    import pandas as pd
    import datetime

    # rede file_name = "N:\LATAM-SaoPaulo\Department\IS\Aniver.xlsx"
    file_name = "Aniver.xlsx"

    # Lê o arquivo .xlsx em um DataFrame do Pandas
    df = pd.read_excel(file_name)

    # Converte campo para tipo date
    df['Dt-Nasc'] = pd.to_datetime(df['Dt-Nasc'], format='%d/%m/%Y')

    # Carrega data de hoje dia e mes
    hoje = datetime.date.today()
    dia = hoje.day
    mes = hoje.month

    # Calcula idade
    df['Idade'] = hoje.year - df['Dt-Nasc'].dt.year

    # Filtra as linhas com aniversários hoje
    aniversarios_hoje = df[(df['Dt-Nasc'].dt.day == dia) & (df['Dt-Nasc'].dt.month == mes)]

    if not aniversarios_hoje.empty:
        enviaemail(aniversarios_hoje)

if __name__ == "__main__":
    main()