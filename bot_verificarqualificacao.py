import logging
logging.basicConfig(level=logging.INFO, filename="log.txt", format="%(asctime)s / %(levelname)s / %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')

from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin


# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def action(self, execution=None):
        import pandas as pd
        base = pd.read_excel(r'C:\Users\rafael\Downloads\CORREÇÃO DAS ORDEM DOS ATOS DO 1º LOTE.xlsx','teste', keep_default_na=False)
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\CORREÇÃO DAS ORDEM DOS ATOS DO 1º LOTE.xlsx").set_nan_as(value='')
        planilha.set_active_sheet('teste')

        dados = planilha.as_list()[1:]


        for index, dados in enumerate(dados, start=2):
            op1 = dados[5]  # operacao de averbacao
            mat = dados[2]
            quali = dados[8]  # qualificacao


            if op1 == 'Abertura de matrícula' and quali != 'Proprietario':
                print(f'{mat} - {op1} e {quali} - errado')

            elif op1 == 'Venda e Compra' and (quali != 'Transmitente' or 'Adquirente'):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Averbação de Construção' and (quali != 'Interessado'):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Aforamento' and (quali != 'Transmitente' or 'Adquirente'):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Adjudicação' and (quali != 'Transmitente' or 'Adquirente'):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Hipoteca' and (quali != 'Devedor' or 'Credor' or 'Anuente' or "Interessado"):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Cancelamento de Hipoteca' and (quali != 'Cancelado' or 'Anuente' or "Interessado"):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Cédula de Crédito Industrial' and (quali != 'Devedor' or 'Credor' or 'Anuente' or "Interessado"):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Doação' and (quali != 'Transmitente' or 'Adquirente'):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Cédula de Crédito Rural e Hipotecária' and (quali != 'Devedor' or 'Credor' or 'Anuente' or "Interessado"):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Mandado de Penhora' and (quali != 'Exequente' or 'Executado' or 'Anuente' or "Interessado"):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            elif op1 == 'Mandado de Indisponibilidade de Bens' and (quali != 'Exequente' or 'Executado' or 'Anuente' or "Interessado"):
                print(f'{mat} - {op1} e {quali} - errado')
                logging.info(f'{mat} - {op1} e {quali} - errado')
            else:
                print(f'{mat} - {op1} e {quali} - OK')
                logging.info(f'{mat} - {op1} e {quali} - OK')






    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
