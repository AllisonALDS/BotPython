import pandas as pd
import time
import smtplib

def enviar_email(esteira, estado, valor):
    """Envia um e-mail com as informações do estado da esteira."""
    remetente = "alvesallison42@gmail.com"
    senha = "quss eqgi bqga iwjw"  # Substitua pela senha do app
    destinatario = "alvesallison42@gmail.com"  # Substitua pelo e-mail do destinatário

    assunto = f"Status da {esteira}"
    mensagem = f"Estado: {estado}\nValor: {valor}"

    email_texto = f"Subject: {assunto}\n\n{mensagem}"

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(remetente, senha)
            server.sendmail(remetente, destinatario, email_texto)
        print(f"E-mail enviado com sucesso: Esteira={esteira}, Estado={estado}, Valor={valor}")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def relatorio(esteira, estado, valor):
    """Cria um relatório e salva no arquivo Excel."""
    novo_dado = pd.DataFrame({
        "esteira": [esteira],
        "valor": [valor],
        "estado": [estado]
    })

    arquivo_relatorio = "Relatorio.xlsx"

    try:
        with pd.ExcelWriter(arquivo_relatorio, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row
            novo_dado.to_excel(writer, index=False, header=False, sheet_name='Sheet1', startrow=startrow)
        print(f"Relatório atualizado: Esteira={esteira}, Valor={valor}, Estado={estado}")
    except FileNotFoundError:
        novo_dado.to_excel(arquivo_relatorio, index=False, sheet_name='Sheet1')
        print(f"Novo relatório criado: Esteira={esteira}, Valor={valor}, Estado={estado}")

def checar_valor(esteira, valor):
    """Verifica o estado da esteira com base no valor e gera um relatório."""
    if valor == 1:
        estado = "Estoque baixo"
    elif valor == 2:
        estado = "Estoque médio"
    elif valor == 3:
        estado = "Estoque cheio"
    else:
        estado = "Valor inválido"

    print(f"{esteira}: {valor} - {estado}")

    if estado != "Valor inválido":
        relatorio(esteira, estado, valor)
        enviar_email(esteira, estado, valor)

def ler_linhas(esteira1, esteira2, esteira3):
    """Itera pelas linhas das esteiras e verifica os valores."""
    for valor1, valor2, valor3 in zip(esteira1, esteira2, esteira3):
        print("Processando Esteira1...")
        checar_valor("Esteira1", valor1)
        time.sleep(1)

        print("Processando Esteira2...")
        checar_valor("Esteira2", valor2)
        time.sleep(1)

        print("Processando Esteira3...")
        checar_valor("Esteira3", valor3)
        time.sleep(1)

# Lê os dados do arquivo Excel
df = pd.read_excel("Esp8266_Receiver.xlsx")

# Define as colunas que representam os valores das esteiras
esteira1 = df["value0"]
esteira2 = df["value1"]
esteira3 = df["value2"]

# Inicia o processamento das esteiras
ler_linhas(esteira1, esteira2, esteira3)
