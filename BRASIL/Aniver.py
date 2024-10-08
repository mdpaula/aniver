# Envia e-mail aos aniversariantes
def enviaemailaniver(df_aniver, df_bcc, df_msganiver):

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
    smtp_server = df_msganiver.loc[df_msganiver['Tag'] == 'mail-server', 'Conteudo'].values[0]
    smtp_user = df_msganiver.loc[df_msganiver['Tag'] == 'de', 'Conteudo'].values[0]
    porta = df_msganiver.loc[df_msganiver['Tag'] == 'porta', 'Conteudo'].values[0]
    server = smtplib.SMTP(smtp_server, porta) 
    tobcc = df_bcc['e-Mail'].tolist()

    for i, linha in df_aniver.iterrows():
        msg = MIMEMultipart('related')
        
        toaddr = linha['e-Mail']
        tobccgestor = linha['Gestor']
        msg['Subject'] = df_msganiver.loc[df_msganiver['Tag'] == 'assunto', 'Conteudo'].values[0]
        msg['From'] = smtp_user
        msg['To'] = toaddr
        msg['Bcc'] = ', ' + tobccgestor + ', '.join(tobcc)

        # Escreva seu HTML aqui

        html = f"""\
        <html>
        <head>
            <meta charset="UTF-8">
            <title>{df_msganiver.loc[df_msganiver['Tag'] == 'titulo', 'Conteudo'].values[0]}</title>
            <style>
                body {{
                    font-family: "Bookman Old Style", sans-serif;
                }}
                p {{
                    font-size: 14pt;
                }}
                h1 {{
                    font-size: 22pt;
                    font-weight: bold;
                }}
            </style>
        </head>
        <body>
            <p>{df_msganiver.loc[df_msganiver['Tag'] == 'saudacao', 'Conteudo'].values[0]}</p>
            <div style="text-align: center; border: 3px solid orange; padding: 10 px;">
                <img src="cid:image" alt="Guerbet">
                <h1>{linha['Nome']},</h1>
                <p>{df_msganiver.loc[df_msganiver['Tag'] == 'p1', 'Conteudo'].values[0]}</p>
                <p>{df_msganiver.loc[df_msganiver['Tag'] == 'p2', 'Conteudo'].values[0]}</p>
            </div>
        </body>
        </html>
        """

        # Anexa o HTML à mensagem
        msg.attach(MIMEText(html, 'html'))

        # Carregue a imagem e anexe-a
        img = df_msganiver.loc[df_msganiver['Tag'] == 'img', 'Conteudo'].values[0]
        with open(img, 'rb') as f:
            image_data = f.read()
            image = MIMEImage(image_data)
            image.add_header('Content-ID', '<image>')
            msg.attach(image)

        # Envia o e-mail
        toaddrs = [toaddr] + [tobccgestor] + tobcc
        server.sendmail(smtp_user, toaddrs, msg.as_string())
    
    # Fecha a conexão com o servidor SMTP
    server.quit()

def enviaemailtempo(df_tempo, df_bcc, df_msgtempo):

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
    smtp_server = df_msgtempo.loc[df_msgtempo['Tag'] == 'mail-server', 'Conteudo'].values[0]
    smtp_user = df_msgtempo.loc[df_msgtempo['Tag'] == 'de', 'Conteudo'].values[0]
    porta = df_msgtempo.loc[df_msgtempo['Tag'] == 'porta', 'Conteudo'].values[0]
    server = smtplib.SMTP(smtp_server, porta)  
    tobcc = df_bcc['e-Mail'].tolist()

    for i, linha in df_tempo.iterrows():
        msg = MIMEMultipart('related')
        
        toaddr = linha['e-Mail']
        tobccgestor = linha['Gestor']
        msg['Subject'] = df_msgtempo.loc[df_msgtempo['Tag'] == 'assunto', 'Conteudo'].values[0]
        msg['From'] = smtp_user
        msg['To'] = toaddr
        msg['Bcc'] = ', ' + tobccgestor + ', '.join(tobcc)

        # Escreva seu HTML aqui

        html = f"""\
        <html>
        <head>
            <meta charset="UTF-8">
            <title>{df_msgtempo.loc[df_msgtempo['Tag'] == 'titulo', 'Conteudo'].values[0]}</title>
            <style>
                body {{
                    font-family: "Bookman Old Style", sans-serif;
                }}
                p {{
                    font-size: 14pt;
                }}
                h1 {{
                    font-size: 22pt;
                    font-weight: bold;
                }}
            </style>
        </head>
        <body>
            <p>{df_msgtempo.loc[df_msgtempo['Tag'] == 'saudacao', 'Conteudo'].values[0]}</p>
            <div style="text-align: center; border: 3px solid orange; padding: 10 px;">
                <img src="cid:image" alt="Guerbet">
                <h1>{linha['Nome']},</h1>
                <p>{df_msgtempo.loc[df_msgtempo['Tag'] == 'p1', 'Conteudo'].values[0]} {linha['Tempo']} {df_msgtempo.loc[df_msgtempo['Tag'] == 'p2', 'Conteudo'].values[0]}</p>
                <p>{df_msgtempo.loc[df_msgtempo['Tag'] == 'p3', 'Conteudo'].values[0]}</p>
            </div>
        </body>
        </html>
        """

        # Anexa o HTML à mensagem
        msg.attach(MIMEText(html, 'html'))

        # Carregue a imagem e anexe-a
        img = df_msgtempo.loc[df_msgtempo['Tag'] == 'img', 'Conteudo'].values[0]
        with open(img, 'rb') as f:
            image_data = f.read()
            image = MIMEImage(image_data)
            image.add_header('Content-ID', '<image>')
            msg.attach(image)

        # Envia o e-mail
        toaddrs = [toaddr] + [tobccgestor] + tobcc
        server.sendmail(smtp_user, toaddrs, msg.as_string())
    
    # Fecha a conexão com o servidor SMTP
    server.quit()

# Inicia programa principal
def main():

    import pandas as pd
    import datetime

    # rede file_name = "N:\LATAM-SaoPaulo\Department\IS\Aniver.xlsx"
    file_name = "Aniver.xlsx"

    # Lê a tab Base do Excel no dataframe dfBase
    df_base = pd.read_excel(file_name, sheet_name='Base')

    df_bcc = pd.read_excel(file_name, sheet_name='Bcc')

    df_msganiver = pd.read_excel(file_name, sheet_name='MsgAniver')

    df_msgtempo = pd.read_excel(file_name, sheet_name='MsgTempo')

    # Converte campo para tipo date
    df_base['Dt-Nasc'] = pd.to_datetime(df_base['Dt-Nasc'], format='%d/%m/%Y')

    # Converte campo para tipo date
    df_base['Dt-Admissao'] = pd.to_datetime(df_base['Dt-Admissao'], format='%d/%m/%Y')

    # Carrega data de hoje dia e mes
    hoje = datetime.date.today()
    dia = hoje.day
    mes = hoje.month

    # Calcula tempo
    df_base['Tempo'] = hoje.year - df_base['Dt-Admissao'].dt.year

    # Calcula jubilado multiplo de 5 anos
    df_base['Tempo5x'] = df_base['Tempo'] % 5

        # Filtra as linhas com aniversários hoje
    df_aniver = df_base[(df_base['Dt-Nasc'].dt.day == dia) & (df_base['Dt-Nasc'].dt.month == mes)]
    
    if not df_aniver.empty and df_msganiver.loc[df_msganiver['Tag'] == 'ativa', 'Conteudo'].values[0]:
        enviaemailaniver(df_aniver, df_bcc, df_msganiver)

    df_tempo = df_base[(df_base['Dt-Admissao'].dt.day == dia) & (df_base['Dt-Admissao'].dt.month == mes) & (df_base['Tempo5x'] == 0)]
                       
    if not df_tempo.empty and df_msgtempo.loc[df_msgtempo['Tag'] == 'ativa', 'Conteudo'].values[0]:
        enviaemailtempo(df_tempo, df_bcc, df_msgtempo)

if __name__ == "__main__":
    main()