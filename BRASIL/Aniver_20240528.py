# Envia e-mail aos aniversariantes
def enviaemail(df_aniver):

    # from email.mime.multipart import MIMEMultipart
    # from email.mime.text import MIMEText
    # from smtplib import SMTP

    # host = 'smtp.gmail.com'
    # port = 587
    # smtp_user = 'marc.mdpaula@gmail.com'
    # password = 'jrnr pwgu zqhe jlf'
    # server = SMTP(host, port)
    # server.starttls()
    # server.login(smtp_user, password)
    
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.image import MIMEImage
    
    # Conecta ao servidor SMTP    
    smtp_server = 'mail.oneguerbet-group.com'
    smtp_user = 'aniversariante@guerbet.com'
    server = smtplib.SMTP(smtp_server, 25)  

    for i, linha in df_aniver.iterrows():
        msg = MIMEMultipart('related')
        msg['Subject'] = 'Feliz Aniversário!'
        msg['From'] = smtp_user
        msg['To'] = linha['e-Mail']

        # Escreva seu HTML aqui
                
        html = f"""\
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Feliz Aniversário!</title>
            <style>
                body {{
                    font-family: "Bookman Old Style", sans-serif;
                }}
                p {{
                    font-size: 14pt;
                }}
                h1 {{
                    font-size: 22pt;
                    color: orange;
                }}
            </style>
        </head>
        <body>
            <p>Olá, bom dia!</p>
            <h1>Feliz Aniversário {linha['Nome']}!</h1>
            <p>Nossos Parabéns por seus {linha['Idade']} aninhos!</p>
            <p>Divirta-se neste seu dia, você merece!</p>
            <img src="cid:image" alt="Guerbet">
        </body>
        </html>
        """
        
        # Anexa o HTML à mensagem
        msg.attach(MIMEText(html, 'html'))

        # Carregue a imagem e anexe-a
        with open('img/Aniver.jpg', 'rb') as f:
            image_data = f.read()
            image = MIMEImage(image_data)
            image.add_header('Content-ID', '<image>')
            msg.attach(image)
            
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