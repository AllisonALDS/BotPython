import pandas as pd
import time
import smtplib
import os

def enviar_email(esteira, estado, valor, date, time_):
    """Envia um e-mail com as informações do estado da esteira."""
    remetente = "alvesallison42@gmail.com"
    senha = "quss eqgi bqga iwjw"  # Substitua pela senha do app
    destinatario = "alvesallison42@gmail.com"  # Substitua pelo e-mail do destinatário

    assunto = f"Status da {esteira}"
    mensagem = f"Estado: {estado}\nValor: {valor}\nData: {date}\nHora: {time_}"

    email_texto = f"Subject: {assunto}\n\n{mensagem}"

    # Envia o e-mail
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, email_texto)
    print(f"E-mail enviado com sucesso: Esteira={esteira}, Estado={estado}, Valor={valor}, Data={date}, Hora={time_}")

def relatorio(esteira, estado, valor, date, time_):
    """Cria um relatório e salva no arquivo Excel com as novas colunas de data e hora."""
    novo_dado = pd.DataFrame({
        "esteira": [esteira],
        "valor": [valor],
        "estado": [estado],
        "data": [date],
        "hora": [time_]
    })

    arquivo_relatorio = "Relatorio.xlsx"

    # Verifica se o arquivo já existe. Se não, cria um novo arquivo.
    if not os.path.exists(arquivo_relatorio):
        # Se o arquivo não existir, cria o arquivo e escreve o cabeçalho
        novo_dado.to_excel(arquivo_relatorio, index=False, sheet_name='Sheet1')
        print(f"Novo relatório criado: Esteira={esteira}, Valor={valor}, Estado={estado}, Data={date}, Hora={time_}")
    else:
        # Se o arquivo já existir, abre e adiciona os novos dados
        with pd.ExcelWriter(arquivo_relatorio, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row
            novo_dado.to_excel(writer, index=False, header=False, sheet_name='Sheet1', startrow=startrow)
        print(f"Relatório atualizado: Esteira={esteira}, Valor={valor}, Estado={estado}, Data={date}, Hora={time_}")

def checar_valor(esteira, valor, date, time_):
    """Verifica o estado da esteira com base no valor e gera um relatório, enviando e-mail."""
    if valor == 1:
        estado = "Estoque baixo"
    elif valor == 2:
        estado = "Estoque médio"
    elif valor == 3:
        estado = "Estoque cheio"
    else:
        estado = "Valor inválido"

    print(f"{esteira}: {valor} - {estado} - Data: {date} - Hora: {time_}")

    if estado != "Valor inválido":
        relatorio(esteira, estado, valor, date, time_)
        enviar_email(esteira, estado, valor, date, time_)

def ler_linhas(esteira1, esteira2, esteira3, dates, times):
    """Itera pelas linhas das esteiras e verifica os valores com data e hora."""
    for valor1, valor2, valor3, date, time_ in zip(esteira1, esteira2, esteira3, dates, times):
        print("Processando Esteira1...")
        checar_valor("Esteira1", valor1, date, time_)
        time.sleep(1)

        print("Processando Esteira2...")
        checar_valor("Esteira2", valor2, date, time_)
        time.sleep(1)

        print("Processando Esteira3...")
        checar_valor("Esteira3", valor3, date, time_)
        time.sleep(1)

# Lê os dados do arquivo Excel
df = pd.read_excel("Esp8266_Receiver.xlsx")

# Define as colunas que representam os valores das esteiras e as novas colunas Date e Time
esteira1 = df["value0"]
esteira2 = df["value1"]
esteira3 = df["value2"]
dates = df["Date"]
times = df["Time"]

# Inicia o processamento das esteiras
ler_linhas(esteira1, esteira2, esteira3, dates, times)
