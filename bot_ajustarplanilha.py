from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin
import logging
<<<<<<< HEAD
logging.basicConfig(level=logging.INFO, filename="C:\RPA\PlanilhaExcel\AjustarPlanilha\AjustarPlanilha\logs\log_ajuste_patricia.txt", format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')
=======
logging.basicConfig(level=logging.INFO, filename="C:\RPA\PlanilhaExcel\AjustarPlanilha\AjustarPlanilha\logs\log_ajuste_patricia.csv", format="%(asctime)s $ %(message)s", datefmt='%d/%m/%Y %I:%M:%S %p')
>>>>>>> f61b942b4a492a43a1800bb1970fd04cf2d321ee



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

<<<<<<< HEAD
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\atos_patricia-3.xlsx").set_nan_as(value='')
        planilha.set_active_sheet('Original')
=======
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\Patricia_12.04.2023-22.05.2023.xlsx").set_nan_as(value='')
        planilha.set_active_sheet('Página1')
>>>>>>> f61b942b4a492a43a1800bb1970fd04cf2d321ee


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
<<<<<<< HEAD
        planilha.write(r"C:\Users\rafael\Downloads\atos_patricia-3.xlsx")
        self.wait(10000)
        print('FINALIZOU AJUSTE')
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\atos_patricia-3.xlsx").set_nan_as(
            value='')
        planilha.set_active_sheet('Original')
=======
        planilha.write(r"C:\Users\rafael\Downloads\Patricia_12.04.2023-22.05.2023.xlsx")
        self.wait(10000)
        print('FINALIZOU AJUSTE')
        planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\Patricia_12.04.2023-22.05.2023.xlsx").set_nan_as(
            value='')
        planilha.set_active_sheet('Página1')
>>>>>>> f61b942b4a492a43a1800bb1970fd04cf2d321ee

        dados = planilha.as_list()[1:]
        for index, dados in enumerate(dados, start=2):
            operacao = dados[5]
            operacaostring = str(operacao)
            numeroato = dados[6]
            numero = str(numeroato)
            matricula = dados[2]
<<<<<<< HEAD
            if operacaostring == 'Abertura de matrícula' and numero != '0':
                print('diferente')
                logging.info(f'Matricula {matricula} $ {operacaostring} $ ATO_{numeroato} $ DIFERENTE DE 0')
            else:
                print(f'{operacao}/{numero}')
=======
            if operacaostring == 'Abertura de matrícula' and numeroato != 0:
                logging.info(f'Matricula {matricula} $ {operacaostring} $ ATO_{numeroato} $ DIFERENTE DE 0')
            else:

>>>>>>> f61b942b4a492a43a1800bb1970fd04cf2d321ee
                pass


        print('FINALIZADO - ABRINDO PLANILHA')
<<<<<<< HEAD
        self.execute(r"C:\Users\rafael\Downloads\atos_patricia-3.xlsx")


=======
        self.execute(r"C:\Users\rafael\Downloads\Patricia_12.04.2023-22.05.2023.xlsx")
        '''
        if not self.find("colunaH", matching=0.97, waiting_time=10000):
            self.not_found("colunaH")
        self.click()
        self.wait(1000)
        if not self.find("formatar", matching=0.97, waiting_time=10000):
            self.not_found("formatar")
        self.click()
        self.kb_type(text='Data Abreviada')
        self.enter()
        self.wait(500)
        self.alt_f4()
        self.type_keys(['alt', 'l'])
        '''
>>>>>>> f61b942b4a492a43a1800bb1970fd04cf2d321ee

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()