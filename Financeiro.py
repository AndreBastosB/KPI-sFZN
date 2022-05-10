import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

#AUTOMAÇÃO DOS SEGUINTES KPI'S:
    # - TRANSAÇÕES
    # - REPASSE P/ FARMÁCIAS.
    # - TICKET MÉDIO.
    # - COMISSÃO FARMÁCIAS.
    # - SPREAD.
    # - ENTREGADORES - CLIENTES
    # - ENTREGADORES - SUBSÍDIOS.
    # - TAXA CIELO / PAYU.
    
#OBS: INSERIR NA VARIAVEL 'LINHA' O NUMERO DA LINHA REFERENTE AO PRIMEIRO PEDIDO DO MES DESEJADO.

ano = input ("Insira o ano desejado (Formato AAAA): ")
mes = input ("Insira o mês desejado (Formato MM): ")

planilhaBase = 'Relatorio de Pedidos - Oficiosa (2020)'
planilhaEditada = 'Teste1'

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

#Seleciona o arquivo que estarei visualizando / editando.
sheet = client.open(planilhaBase).get_worksheet(1)
sheet2 = client.open(planilhaEditada).sheet1

linha = input('linha do primeiro pedido do mes na planilha Oficiosa, aba2: ')

quantidadePedidos = 0
transacoes = float(0)
repasseFarmacias = float(0)
ticketMedio = float(0)
comissaoFarmacias = float(0)
spread = float(0)
deliveryCliente = float(0)
subsEntregadores = float(0)
taxaMaquina = float(0)

data = mes + '/' + ano
taxaDjango = 0.039
taxaPayU = 0.0287
payUNome = "PayU"

while data in sheet.cell(linha, 3).value:
    
    #DATA DO PEDIDO
    dataPedido = sheet.cell(linha, 3).value
    
    #RETIRADA DOS VALORES DE TRANSAÇÕES NA PLANILHA (COLUNA D)
    transacoes1 = sheet.cell(linha, 4).value
    transacoes2 = (transacoes1.translate({ord(i): None for i in 'R$ '}))
    
    #If statement caso o valor possua 4 dígitos, if corrige erro de formatação.
    if len(transacoes2) == 8:
        transacoes3 = transacoes2.translate({ord(','): None})
        transacoes4 = float(transacoes3)
    else:
        transacoes3 = transacoes2.replace(',', '.')
        transacoes4 = float(transacoes3)
        
    transacoes += transacoes4
    
    #RETIRADA DOS VALORES DE REPASSE P/ FARMACIAS NA PLANILHA (COLUNA O)
    repasseFarmacias1 = sheet.cell(linha, 15).value
    repasseFarmacias2 = (repasseFarmacias1.translate({ord(i): None for i in 'R$ '}))
    
    #If statement caso o valor possua 4 dígitos, if corrige erro de formatação.
    if len(repasseFarmacias2) == 8:
        repasseFarmacias3 = repasseFarmacias2.translate({ord(','): None})
        repasseFarmacias4 = float(repasseFarmacias3)
    else:
        repasseFarmacias3 = repasseFarmacias2.replace(',', '.')
        repasseFarmacias4 = float(repasseFarmacias3)
        
    repasseFarmacias += repasseFarmacias4
    
    #RETIRADA DOS VALORES DE TED P/ FARMACIAS NA PLANILHA (COLUNA J)
    ticketMedio1 = sheet.cell(linha, 10).value
    ticketMedio2 = (ticketMedio1.translate({ord(i): None for i in 'R$ '}))
    
    #If statement caso o valor possua 4 dígitos, if corrige erro de formatação.
    if len(ticketMedio2) == 8:
        ticketMedio3 = ticketMedio2.translate({ord(','): None})
        ticketMedio4 = float(ticketMedio3)
    else:
        ticketMedio3 = ticketMedio2.replace(',', '.')
        ticketMedio4 = float(ticketMedio3)
    
    ticketMedio += ticketMedio4
    
    #RETIRADA DOS VALORES DE COMISSAO NA PLANILHA (COLUNA N)
    comissaoFarmacias1 = sheet.cell(linha, 14).value
    comissaoFarmacias2 = (comissaoFarmacias1.translate({ord(i): None for i in 'R$ '}))
    comissaoFarmacias3 = comissaoFarmacias2.replace(',', '.')
    comissaoFarmacias4 = float(comissaoFarmacias3)
    comissaoFarmacias += comissaoFarmacias4
    
    #RETIRADA DOS VALORES DE SPREAD NA PLANILHA (COLUNA AE)
    spread1 = sheet.cell(linha, 31).value
    spread2 = (spread1.translate({ord(i): None for i in 'R$ '}))
    spread3 = spread2.replace(',', '.')
    spread4 = float(spread3)
    spread += spread4
    
    #RETIRADA DOS VALORES DE DELIVERY COBRADO NA PLANILHA (COLUNA R)
    deliveryCliente1 = sheet.cell(linha, 18).value
    deliveryCliente2 = (deliveryCliente1.translate({ord(i): None for i in 'R$ '}))
    deliveryCliente3 = deliveryCliente2.replace(',', '.')
    deliveryCliente4 = float(deliveryCliente3)
    deliveryCliente += deliveryCliente4
    
    #RETIRADA DOS VALORES DE DELIVERY CUSTO COBRADO NA PLANILHA (COLUNA S)
    subsEntregadores1 = sheet.cell(linha, 19).value
    subsEntregadores2 = (subsEntregadores1.translate({ord(i): None for i in 'R$ '}))
    subsEntregadores3 = subsEntregadores2.replace(',', '.')
    subsEntregadores4 = float(subsEntregadores3)
    subsEntregadores += subsEntregadores4
    
    if payUNome in sheet.cell(linha, 2).value:
        taxaMaquina1 = transacoes4 * taxaPayU
    else:
        taxaMaquina1 = transacoes4 * taxaDjango
        
    taxaMaquina += taxaMaquina1
    
    linha += 1
    quantidadePedidos += 1
    time.sleep(20)
    

#TOTAL DE TRANSAÇOES (COLUNA D PLANILHA OFICIOSA, SHEET 2)
transacoesFinal = "{:.2f}".format(transacoes)

#TOTAL DE REPASSE P/ FARMACIAS (COLUNA O PLANILHA OFICIOSA, SHEET 2)
repasseFarmaciasFinal = "{:.2f}".format(repasseFarmacias)

#TOTAL DE VENDAS DIVIDIDO PELO NUMERO DE PEDIDOS + FORMATAÇÃO (COLUNDA J PLANILHA OFICIOSA, SHEET 2)
ticketMedioFinal1 = ticketMedio / quantidadePedidos
ticketMedioFinal = "{:.2f}".format(ticketMedioFinal1)

#TOTAL DE TED P/ FARMACIAS (COLUNA N PLANILHA OFICIOSA, SHEET 2)
comissaoFarmaciasFinal = "{:.2f}".format(comissaoFarmacias)

#TOTAL DE SPREAD (COLUNA AE PLANILHA OFICIOSA, SHEET 2)
spreadFinal = "{:.2f}".format(spread)

#TOTAL DE DELIVERY COBRADO (COLUNA R PLANILHA OFICIOSA, SHEET 2)
deliveryClienteFinal = "{:.2f}".format(deliveryCliente)

#TOTAL DE CUSTO DELIVERY - DELIVERY COBRADO (COLUNA S-R PLANILHA OFICIOSA, SHEET 2)
subsidioEntregadores1 = subsEntregadores - deliveryCliente
subsidioEntregadores = "{:.2f}".format(subsidioEntregadores1)

#TOTAL DE CUSTO DAS TRANSAÇÕES PELAS OPERADORAS BANCARIAS (COLUNA D -3,9% OU -2,87% (PAYU) PLANILHA OFICIOSA, SHEET 2)
taxaMaquinaFinal = "{:.2f}".format(taxaMaquina)

print ('Transações: R$' + transacoesFinal)
print ('Repasse p/ Farmacias: R$' + repasseFarmaciasFinal)
print ('Ticket Médio: R$' + ticketMedioFinal)
print ('Comissão Farmacias: R$' + comissaoFarmaciasFinal)
print ('Spread: R$' + spreadFinal)
print ('Entregadores - Clientes: R$' + deliveryClienteFinal)
print ('Entregadores - Subsídio: R$' + subsidioEntregadores)
print ('Taxa Cielo / PayU: R$' + taxaMaquinaFinal)

#Linha correspondentes no KPI.
linhasTransacoes = 19
linhaRepasse = 20
linhaTicketMedio = 22
linhaComissao = 23
linhasSpread = 24
linhaEntregCl = 25
linhaEntregSubs = 26
linhaTaxaCielo = 27

#Coluna do mês correspondente.
if ano == "2019":
    Mes = int(mes) - 5
if ano == "2020":
    Mes = int(mes) + 7
if ano == "2021":
    Mes = int(mes) + 19
elif ano == "2022":
    Mes = int(mes) + 31


sheet2.update_cell (linhasTransacoes, Mes, transacoesFinal)
sheet2.update_cell (linhaRepasse, Mes, repasseFarmaciasFinal)
sheet2.update_cell (linhaTicketMedio, Mes, ticketMedioFinal)
sheet2.update_cell (linhaComissao, Mes, comissaoFarmaciasFinal)
sheet2.update_cell (linhasSpread, Mes, spreadFinal)
sheet2.update_cell (linhaEntregCl, Mes, deliveryClienteFinal)
sheet2.update_cell (linhaEntregSubs, Mes, subsidioEntregadores)
sheet2.update_cell (linhaTaxaCielo, Mes, taxaMaquinaFinal)







