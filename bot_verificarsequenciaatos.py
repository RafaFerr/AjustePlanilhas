import logging
logging.basicConfig(level=logging.INFO, filename="log_sequencia.csv", format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')

from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin


# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def action(self, execution=None):
        import pandas as pd
        base = pd.read_excel(r'C:\Users\rafael\Downloads\CORREÇÃO DAS ORDEM DOS ATOS DO 1º LOTE.xlsx','plan1', keep_default_na=False)
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\qualificacao_atos_ancelmo.xlsx").set_nan_as(value='')
        planilha.set_active_sheet('teste')
        planilha2 = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\qualificacao_atos_ancelmo.xlsx").set_nan_as(value='')
        planilha2.set_active_sheet('atos')
        dados = planilha.as_list()[1:]
        dados2 = planilha2.as_list()[2:]
        for i in range(1779):
            matricula = str(base['MATRICULA'][i])
            num1 = str(base['NUMERO DO ATO'][i])
            num2 = str(base['NUMERO DO ATO'][i + 1])

            if num1 == num2:
                logging.info(f'MATRICULA_{matricula} $ ATO_{num1} $ DUPLICADO')
                print(f'{matricula} -$ {num1} $ {num2} $ errado' )
            else:
                print(f'{matricula} - {num1} e {num2} sequencia certa')
                pass













    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
