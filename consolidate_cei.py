import os
import re
import xlrd
import pandas as pd

def consolidate_cei_extracts(base_folder = 'extratos_cei', save_to_file = False):
    cols = ['Data Negócio', 'C/V', 'Mercado', "Prazo", 'Código', 'Especificação do Ativo', 'Quantidade', 'Preço (R$)', 'Valor Total (R$)']

    cei_files = os.listdir(base_folder)
    transactions = pd.DataFrame()
    for cei_file in cei_files:
        if (cei_file == '.DS_Store') or ("~lock" in cei_file): continue # for macOS compatibility and removal of temporary files

        broker = re.search(r'negociacoes_cei_(.*)\.xlsx?$', cei_file)[1]
        filepath = os.path.join(base_folder, cei_file)

        wb = xlrd.open_workbook(filepath, logfile=open(os.devnull, 'w'))
        file_transactions = pd.read_excel(wb, header = 10, usecols=cols, engine='xlrd').dropna(subset = ['Código'])
        file_transactions['Corretora'] = broker

        transactions = transactions.append(file_transactions, ignore_index = True)

    transactions.rename(columns = {
        'Data Negócio': 'Data',
        'C/V': 'Fluxo',
        'Código': 'Codigo',
        'Especificação do Ativo': 'Ativo',
        'Preço (R$)': 'Preco',   
        'Valor Total (R$)': 'Valor Total'
    }, inplace = True)

    transactions['Data'] = pd.to_datetime(transactions['Data'], dayfirst = True)
    transactions['Prazo'] = pd.to_datetime(transactions['Prazo'], dayfirst = True)
    transactions.set_index('Data', inplace = True)
    transactions.sort_index(inplace = True)

    for col in ['Fluxo', 'Mercado', 'Codigo', 'Ativo']:
        transactions[col] = transactions[col].str.strip()
    
    transactions['Codigo'] = transactions['Codigo'].str.replace("F$", "")
    transactions['Quantidade'] = transactions['Quantidade'] * transactions['Fluxo'].map({"C": 1, "V": -1})
    transactions['Valor Total'] = transactions['Valor Total'] * transactions['Fluxo'].map({"C": 1, "V": -1})
    transactions['Tipo'] = transactions.apply(define_transaction_type, axis = 1)


    if save_to_file:
        save_data = transactions.copy()
        save_data.index = save_data.index.date
        save_data.index.name = "Data"
        save_data.reset_index().to_excel('consolidado_cei.xlsx', index = False)
        print('Dados salvos na planilha consolidado_cei.xlsx')

    return transactions


def define_transaction_type(row):
    if ("11" in row["Codigo"].upper() or "12" in row["Codigo"].upper()) and "FII " in row["Ativo"].upper():
        return "FII"
    elif row["Mercado"].lower().strip() in ("opção de compra", "opção de venda") and not pd.isna(row["Prazo"]):
        return "Opção"
    elif row["Mercado"].lower().strip() == "exercicio de opções":
        return "Opção (Exercício)"
    elif row["Mercado"].lower().strip() in ("mercado a vista", "merc. fracionario"):
        return "Ação"
    else:
        return "desconhecido"


if __name__ == "__main__":
    consolidate_cei_extracts(save_to_file = True)