import win32com.client as w32
import re
from datetime import datetime

outlook = w32.Dispatch('Outlook.Application').GetNamespace('MAPI')
account = outlook.Folders['daphiny.pinheiro@legalcontrol.com.br']
inbox = account.Folders['Caixa de Entrada']

#Pega no nome dos Clientes
def extrair_nome_cliente(corpo_email):
    padrao = r'Cliente:\s*(.*)'
    match = re.search(padrao, corpo_email)
    if match:
        return match.group(1).strip()
    return "Cliente Desconhecido"

#Pega o numero das Solicitações e Envios
def extrair_notificacoes(corpo_email):
    notificacoes = {
        'solicitacoes': 0,
        'enviadas': 0
    }
    
    padroes = [
        r'Total\s+de\s+solicitações\s+\w+:\s*(\d+)', 
        r'enviada\(s\):\s*(\d+)'  
    ]
    
    
    for padrao in padroes:
        match = re.search(padrao, corpo_email)
        if match:
            notificacoes['solicitacoes'] += int(match.group(1))
            if 'enviada' in padrao:
                notificacoes['enviadas'] += int(match.group(1))
    
    return notificacoes


data_atual = datetime.now().date()
print(f"Data atual: {data_atual}")

#Filtra os emails do dia de Hoje e do no-reply@lcontrol.com.br
def filtrar_emails_dia_atual(items, remetente):
    emails_do_dia = []
    for item in items:
        data_recebimento = item.ReceivedTime
        if data_recebimento.date() == data_atual and item.SenderEmailAddress == remetente:
            emails_do_dia.append(item)
    return emails_do_dia

remetente_especifico = "no-reply@lcontrol.com.br"

print(f"Filtrando emails do {remetente_especifico}")
emails = filtrar_emails_dia_atual(inbox.Items, remetente_especifico)
print(f"Número de emails encontrados: {len(emails)}")


emails_por_cliente = {}


for m in emails:
    
    corpo_email = m.Body
   
    nome_cliente = extrair_nome_cliente(corpo_email)
    print(f"Processando email para o cliente: {nome_cliente}")
 
    notificacoes = extrair_notificacoes(corpo_email)
    
    email_id = m.EntryID
    
    if nome_cliente not in emails_por_cliente:
        emails_por_cliente[nome_cliente] = []
    

    emails_por_cliente[nome_cliente].append({
        'solicitacoes': notificacoes['solicitacoes'],
        'enviadas': notificacoes['enviadas']
    })

clientes_ordenados = sorted(emails_por_cliente.keys(), key=str.casefold)


print("Listagem de emails por cliente:")
for cliente in clientes_ordenados:
    print("===============================================================")
    print(f"Cliente: {cliente}")
    for i, email in enumerate(emails_por_cliente[cliente], start=1):
        num_solicitacoes = email['solicitacoes']
        num_enviadas = email['enviadas']

        if num_solicitacoes == 0 and num_enviadas == 0:
            print("OK")
        else:
            print("-----------------------")
            print(f"| Solicitações: {num_solicitacoes}     |")
            print(f"| Envios: {num_enviadas}           |")
            print("-----------------------")
