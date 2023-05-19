import logging
logging.basicConfig(level=logging.INFO, filename="log.txt", format="%(asctime)s / %(levelname)s / %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')

# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False
from botcity.plugins.excel import BotExcelPlugin


def main():


    planilha = BotExcelPlugin().read(r"C:\Users\Rafael\Downloads\CORREÇÃO DAS ORDEM DOS ATOS DO 1º LOTE.xlsx").set_nan_as(value='')
    planilha.set_active_sheet('plan2')

    dados = planilha.as_list()[1:]

    for index, dados in enumerate(dados, start=2):
        op1 = dados[5]  # operacao de averbacao
        mat = dados[2]
        quali = dados[8]  # qualificacao

        if op1 == 'Abertura de matrícula' and quali == 'Proprietario':
            print(f'{mat} - {op1} e {quali} - Certo')
        elif op1 == 'Venda e Compra' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Averbação de Construção' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Aforamento' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Adjudicação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Hipoteca' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado" or quali == "Fiador":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cancelamento de Hipoteca' and quali == 'Cancelado' or quali == 'Anuente' or quali == "Interessado":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cédula de Crédito Industrial' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Doação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cédula de Crédito Rural e Hipotecária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Mandado de Penhora' and quali == 'Exequente' or quali == 'Executado' or quali == 'Anuente' or quali == "Interessado":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Mandado de Indisponibilidade de Bens' and quali == 'Exequente' or quali == 'Executado' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Desmembramento' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Aditivo' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Alienação Fiduciária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Arrendamento' and quali == 'Arrendante' or quali == 'Arrendatario' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Averbação' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Averbação de Óbito' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cancelamento de Aforamento' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cancelamento de Matrícula' and quali == 'Cancelado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cancelamento de Ônus' and quali == 'Cancelado' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cancelamento de Penhora' and quali == 'Cancelado' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Carta de Arrematação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Caução' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cédula de Crédito Comercial' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cédula Rural Pignoratícia e Hipotecária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Cessão de Direitos' and quali == 'Cedente' or quali == 'Cessionario':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Demolição' and (quali == 'Interessado'):
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Desapropriação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Erro Evidente' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Formal de Partilha' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Liberação de Hipoteca' and quali == 'Cancelado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Pagamento' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Permuta' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Re-Ratificação' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Reserva de Vegetação' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Retificação de Área' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Retificação Nome' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Unificação e Encerramento' and quali == 'Interessado':
            print(f'{mat} - {op1} e {quali} - Certo')
            
        elif op1 == 'Venda e Compra com Hipoteca' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Credor' or quali == 'Devedor' or quali == 'Anuente':
            print(f'{mat} - {op1} e {quali} - Certo')

        else:
            print(f'{mat} - {op1} e {quali} - VERIFICAR')
            logging.info(f'{mat} - {op1} e {quali} - VERIFICAR')


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
