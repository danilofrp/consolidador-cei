import os
import re
import xlrd
import pandas as pd

def consolidate_cei_extracts(base_folder = 'extratos_cei', save_to_file = False):
    cols = ['Data Negócio', 'C/V', 'Mercado', 'Código', 'Especificação do Ativo', 'Quantidade', 'Preço (R$)', 'Valor Total (R$)']

    cei_files = os.listdir(base_folder)
    transactions = pd.DataFrame()
    for cei_file in cei_files:
        if (cei_file == '.DS_Store'): continue # for macOS compatibility

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
    transactions.set_index('Data', inplace = True)
    transactions.sort_index(inplace = True)

    for col in ['Fluxo', 'Mercado', 'Codigo', 'Ativo']:
        transactions[col] = transactions[col].str.strip()
    
    transactions['Codigo'] = transactions['Codigo'].apply(lambda s: re.sub('F$', '', s))

    if save_to_file:
        save_data = transactions.copy()
        save_data.index = save_data.index.date
        save_data.to_excel('consolidado_cei.xlsx')
        print('Dados salvos na planilha consolidado_cei.xlsx')

    return transactions


if __name__ == "__main__":
    consolidate_cei_extracts(save_to_file = True)