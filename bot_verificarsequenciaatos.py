import logging
logging.basicConfig(level=logging.INFO, filename="C:\RPA\PlanilhaExcel\AjustarPlanilha\AjustarPlanilha\logs\log_sequencia_patricia.txt", format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')

from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin




class Bot(DesktopBot):
    def action(self, execution=None):
        import pandas as pd
        base = pd.read_excel(r"C:\Users\rafael\Downloads\atos_patricia-3.xlsx",'Atos', keep_default_na=False)

        for i in range(3415):
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
