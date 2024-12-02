import pandas as pd
import time

# Lê os dados do arquivo Excel
df = pd.read_excel("Esp8266_Receiver.xlsx")

# Define as colunas do Excel
esteira1 = df["value0"]
esteira2 = df["value1"]
esteira3 = df["value2"]
datas = df["Date"]
horarios = df["Time"]

def relatorio(esteira, estado, valor, data, hora):
    """Salva os dados em um arquivo Excel."""
    novo_dado = pd.DataFrame({
        "Data": [data],
        "Hora": [hora],
        "Esteira": [esteira],
        "Valor": [valor],
        "Estado": [estado]
    })
    
    # Salva ou cria o arquivo Excel
    try:
        with pd.ExcelWriter("Relatorio.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            startrow = writer.sheets["Sheet1"].max_row  # Próxima linha disponível
            novo_dado.to_excel(writer, index=False, header=False, sheet_name="Sheet1", startrow=startrow)
    except FileNotFoundError:
        novo_dado.to_excel("Relatorio.xlsx", index=False, sheet_name="Sheet1")

def checar_valor(esteira, valor, data, hora):
    """Verifica o valor e registra o estado no relatório."""
    estado = {
        1: "Estoque baixo",
        2: "Estoque médio",
        3: "Estoque cheio"
    }.get(valor, "Valor inválido")
    
    print(f"{data} {hora} - {esteira}: {valor} - {estado}")
    if estado != "Valor inválido":
        relatorio(esteira, estado, valor, data, hora)

def ler_linhas(esteira1, esteira2, esteira3, datas, horarios):
    """Itera pelos valores das esteiras, datas e horários."""
    for valor1, valor2, valor3, data, hora in zip(esteira1, esteira2, esteira3, datas, horarios):
        checar_valor("Esteira1", valor1, data, hora)
        time.sleep(1)

        checar_valor("Esteira2", valor2, data, hora)
        time.sleep(1)

        checar_valor("Esteira3", valor3, data, hora)
        time.sleep(1)

# Inicia o processamento
ler_linhas(esteira1, esteira2, esteira3, datas, horarios)
