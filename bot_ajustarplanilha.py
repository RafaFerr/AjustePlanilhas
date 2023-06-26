from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin
import logging

logging.basicConfig(level=logging.INFO, filename="C:\RPA\PlanilhaExcel\AjustarPlanilha\AjustarPlanilha\logs\log_ajuste_ancelmo.txt", format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')

'''
A PLANILHA DE QUALIFICAÇÃO DEVERA TER A SEGUINTE ORDERM DAS COLUNAS
Carimbo de data/hora,  E-mail, MATRICULA, NUMERO PROTOCOLO, TIPOATO, OPERACAO, NUMERO DO ATO, DATA, QUALIFICACAO, CPF/CNPJ, NOME, ESTADO CIVIL;

A PLANILHA COM OS ATOS DEVEM TER A SEQUENCIA ABAIXO DE COLUNAS:
Carimbo de data/hora,  E-mail, MATRICULA, NUMERO PROTOCOLO, TIPOATO, OPERACAO, NUMERO DO ATO, DATA.
FORMULAS DO EXCEL PARA ATO_COMPLETO
TIPO1: =SES(E2="ABERTURA";"M.";E2="AVERBAÇÃO";"Av.";E2="Registro";"R.")
ATO_COMPLETO: =CONCAT(I2;G2)
TIPO2: =SES(E2="ABERTURA";"M.";E2="AVERBAÇÃO";"R.";E2="Registro";"Av.")
ATO_COMPLETO2: =CONCAT(K2;G2)

A PLANILHA COM A QUALIFICACAO DEVEM TER A SEQUENCIA ABAIXO DE COLUNAS:
Carimbo de data/hora,  E-mail, MATRICULA,NUMERO DO ATO,QUALIFICACAO, CPF/CNPJ, NOME, ESTADO CIVIL;

'''

class Bot(DesktopBot):
    def action(self, execution=None):


        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\atos_ancelmo.xlsx").set_nan_as(value='')
        planilha.set_active_sheet('Ancelmo')

        dados = planilha.as_list()[1:]
        for index, dados in enumerate(dados, start=2):

            op1 = dados[8]#operacao de averbacao
            num1 = dados[9]#numero de ato da averbacao
            dt1 = dados[10]#data do ato de averbacao
            op2 = dados[11]#operacao de registro
            num2 = dados[12]#numero de ato de registro
            dt2 = dados[13]#data do ato de registro

            print(index)
            if op1 == '' and op2 != '':
                planilha.set_cell('F', index, op2)
                planilha.set_cell('G', index, num2)
                planilha.set_cell('H', index, dt2)
            elif op1 != '' and op2 == '':
                planilha.set_cell('F', index, op1)
                planilha.set_cell('G', index, num1)
                planilha.set_cell('H', index, dt1)
            else:
                continue
        self.wait(2000)
        planilha.remove_columns(['I', 'J', 'K', 'L', 'M', 'N'])

        planilha.write(r"C:\Users\rafael\Downloads\atos_ancelmo.xlsx")
        self.wait(10000)
        print('FINALIZOU AJUSTE')
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\atos_ancelmo.xlsx").set_nan_as(value='')
        planilha.set_active_sheet('Ancelmo')


        dados = planilha.as_list()[1:]
        for index, dados in enumerate(dados, start=2):
            operacao = dados[5]
            operacaostring = str(operacao)
            numeroato = dados[6]
            numero = str(numeroato)
            matricula = dados[2]

            if operacaostring == 'Abertura de matrícula' and numero != '0':
                print('diferente')
                logging.info(f'Matricula {matricula} $ {operacaostring} $ ATO_{numeroato} $ DIFERENTE DE 0')
            else:
                pass

        print('FINALIZADO - ABRINDO PLANILHA')

        self.execute(r"C:\Users\rafael\Downloads\atos_ancelmo.xlsx")


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()