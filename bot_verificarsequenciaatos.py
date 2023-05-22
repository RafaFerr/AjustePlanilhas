import logging
logging.basicConfig(level=logging.INFO, filename="C:\RPA\PlanilhaExcel\AjustarPlanilha\AjustarPlanilha\logs\log_sequenciaatos_patricia.csv", format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')

from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin


# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def action(self, execution=None):
        import pandas as pd
        base = pd.read_excel(r"C:\Users\rafael\Downloads\Patricia_12.04.2023-22.05.2023.xlsx",'Atos', keep_default_na=False)

        for i in range(2104):
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
