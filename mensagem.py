import pandas as pd
import pywhatkit as kit
import time

# Carregar arquivo .xlsx
file_name = 'clientes.xlsx'
df = pd.read_excel(file_name)

# Validar colunas
if 'cliente' not in df.columns or 'telefone' not in df.columns:
    raise ValueError("O arquivo precisa conter as colunas 'cliente' e 'telefone'.")

# Carregar a mensagem de arquivo .txt
message_file = 'mensagem.txt'
try:
    with open(message_file, 'r', encoding='utf-8') as file:
        message_template = file.read().strip()
except FileNotFoundError:
    raise FileNotFoundError(f"The file '{message_file}' was not found. Please create it with the message text.")

# Função de envio de mensagens
def send_whatsapp_messages(dataframe, simulate):
    for index, row in dataframe.iterrows():
        cliente = row['cliente']
        telefone = row['telefone']
        
        # Garantir formato
        if not telefone.startswith('+'):
            telefone = '+55' + telefone.strip()  # Adicionar +55 

        # Pular cliente se o formato do telefone for inválido
        if len(telefone) != 18:
            print(f"[IGNORADO] Número de telefone inválido para {row['cliente']}: {telefone}")
            continue
        
        # Substituir {nome_cliente} pelo nome
        message = message_template.replace("{nome_cliente}", cliente)

        if simulate:
            # Printar a mensagem ao invés de mandar
            print(f"[SIMULAÇÃO] Mensagem para {cliente} - {telefone}: {message}")
        else:
            # Enviar mensagem
            try:
                print(f"Sending message to {cliente} ({telefone})...")
                kit.sendwhatmsg_instantly(telefone, message, wait_time=10, tab_close=True)
                time.sleep(10)  # Timer
            except Exception as e:
                print(f"Falha ao enviar mensagem para{cliente} ({telefone}): {e}")
    
    # Final da execução
    print("\nTodas as mensagens foram enviadas!")
    
    

# Input para simulação
while True:
    user_input = input("Rodar em modo simulação? s/n ").strip().lower()
    if user_input in ['s', 'n']:
        simulate = user_input == 's'
        break
    else:
        print("Entrada inválida. Digite 's' ou 'n'")

# Chamar a função
send_whatsapp_messages(df, simulate=simulate)