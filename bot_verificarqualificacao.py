import logging
logging.basicConfig(level=logging.INFO, filename="C:\RPA\PlanilhaExcel\AjustarPlanilha\AjustarPlanilha\logs\log_qualificacao_patricia.txt", format="%(asctime)s $ %(message)s",datefmt='%d/%m/%Y %I:%M:%S %p')

# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False
from botcity.plugins.excel import BotExcelPlugin


def main():
    planilha = BotExcelPlugin().read(r"C:\Users\rafael\Downloads\atos_patricia-3.xlsx").set_nan_as(value='')
    planilha.set_active_sheet('Original')

    dados = planilha.as_list()[1:]

    for index, dados in enumerate(dados, start=2):
        op1 = dados[5]  # operacao de averbacao
        mat = dados[2]
        num_ato = dados[6]
        quali = dados[8]  # qualificacao

        if op1 == 'Abertura de matrícula' and quali == 'Proprietario':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')
        elif op1 == 'Venda e Compra' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Averbação de Construção' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Aforamento' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Adjudicação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Hipoteca' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado" or quali == "Fiador":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Hipoteca' and quali == 'Cancelado' or quali == 'Anuente' or quali == "Interessado":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cédula de Crédito Industrial' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Doação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cédula de Crédito Rural e Hipotecária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Mandado de Penhora' and quali == 'Exequente' or quali == 'Executado' or quali == 'Anuente' or quali == "Interessado":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Mandado de Indisponibilidade de Bens' and quali == 'Exequente' or quali == 'Executado' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Desmembramento' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Aditivo' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Alienação Fiduciária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Arrendamento' and quali == 'Arrendante' or quali == 'Arrendatario' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Averbação' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Averbação de Óbito' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Aforamento' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Matrícula' and quali == 'Cancelado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Ônus' and quali == 'Cancelado' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Penhora' and quali == 'Cancelado' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Carta de Arrematação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Caução' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cédula de Crédito Comercial' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cédula Rural Pignoratícia e Hipotecária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado":
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cessão de Direitos' and quali == 'Cedente' or quali == 'Cessionario':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Demolição' and (quali == 'Interessado'):
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Desapropriação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Erro Evidente' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Formal de Partilha' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Liberação de Hipoteca' and quali == 'Cancelado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Pagamento' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Permuta' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Re$Ratificação' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Reserva de Vegetação' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Retificação de Área' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Retificação Nome' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Unificação e Encerramento' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra com Hipoteca' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Credor' or quali == 'Devedor' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Alteração de Estado Civil' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Arrematação em Hasta Pública' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Arrendamento' and quali == 'Arrendatario' or quali == 'Arrendante':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Arrolamento de Bens' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cadastro Municipal' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento da Propriedade Fiduciária' and quali == 'Anuente' or quali == 'Cancelado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Averbação' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cancelamento de Registro':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')
            logging.info(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ TERA QUE SER CANCELADO NO REGISTER, NÃO TEM QUALIFICAÇÃO PARA ESTA OPERAÇÃO')

        elif op1 == 'Cancelamento de Usufruto' and quali == 'Interessado' or quali == 'Anuente' or quali == 'Cancelado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cédula de Crédito Banc/Com e Hipotecária' and quali == 'Devedor' or quali == 'Credor' or quali == 'Anuente' or quali == "Interessado" or quali == 'Avalista':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Cessão de Direitos de Compromisso' and quali == 'Cedente' or quali == 'Cessionario':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Composição e Confissão de Dívidas' and quali == 'Credor' or quali == 'Devedor' or quali == 'Interveniente' or quali == 'Fiador':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Compromisso de Venda e Compra' and quali == 'Compromitente Vendedor' or quali == 'Compromitente Comprador':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Contrato de Locação' and quali == 'Locador' or quali == 'Locatario':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Dação em Pagamento' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Divisão Amigável' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Doação da Nua Propriedade' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Encerramento de Matrícula' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Incorporação de Bens' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Inventário Extrajudicial':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')
            logging.info(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VER ESSA OPERACAO SE TEM QUE OCORRER TRANSFERENCIA DE PROPRIEDADE')

        elif op1 == 'Liberação de Penhora' and quali == 'Cancelado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Locação' and quali == 'Locador' or quali == 'Locatario':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR POIS EXISTE UMA OPERAÇÃO DE CONTRATO DE LOCAÇÃO')
            logging.info((f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR POIS EXISTE UMA OPERAÇÃO DE CONTRATO DE LOCAÇÃO'))

        elif op1 == 'Mandado de Cancelamento de Reg./Av.':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR POIS EXISTE UMA OPERAÇÃO DE CANCELAMENTO DE REGISTRO')
            logging.info((f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR POIS EXISTE UMA OPERAÇÃO DE CANCELAMENTO DE REGISTRO'))

        elif op1 == 'Mandado de Sequestro de Bens' and quali == 'Credor' or quali == 'Devedor' or quali == 'Requerido' or quali == 'Requerente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Mandado de Usucapião' and quali == 'Adquirente' or quali == 'Transmitente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Mandado Judicial' and quali == 'Interessado' or quali == 'Exequente' or quali == 'Executado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Mandado Judicial' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Rescisão de Contrato':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')
            logging.info(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR QUE CONTRATO ESTA SENDO CANCELADO PARA VER SE TEM ALGUM ONUS')

        elif op1 == 'Retificação de Registro' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Servidão de Passagem' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Título de Concessão de Domínio' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Título Definitivo' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Transferência' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Transporte' and quali == 'Interessado' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Unificação (Ao abrir Matrícula)' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Unificação e Encerramento' and quali == 'Interessado':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra com Pacto Comissório' and quali == 'Transmitente' or quali == 'Adquirente' or quali == 'Credor' or quali == 'Devedor' or quali == 'Anuente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra com Reserva de Usufruto' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra com Reserva de Vegetação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra com Reserva de Vegetação' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra da Nua Propriedade' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra da Nua Propriedade e do Usufruto' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra do Usufruto' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        elif op1 == 'Venda e Compra em Carater Pró-Soluto' and quali == 'Transmitente' or quali == 'Adquirente':
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ Certo')

        else:
            print(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR')
            logging.info(f'MATRICULA_{mat} $ {op1} $ NUM_ATO_{num_ato} $ QUALIFICACAO_{quali} $ VERIFICAR')


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()