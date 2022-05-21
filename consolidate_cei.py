import os
import re
import xlrd
import pandas as pd
from unidecode import unidecode

def consolidate_cei_extracts(base_folder = 'movimentacoes', save_to_file = False):
    cols = ['Entrada/Saída','Data','Movimentação','Produto','Instituição','Quantidade','Preço unitário','Valor da Operação']

    cei_files = os.listdir(base_folder)
    
    cei_files = [f for f in cei_files if "~$" not in f]
    
    transactions = pd.DataFrame()
    for cei_file in cei_files:
        if (cei_file == '.DS_Store') or ("~lock" in cei_file): continue # for macOS compatibility and removal of temporary files

        # broker = re.search(r'negociacoes_cei_(.*)\.xlsx?$', cei_file)[1]
        filepath = os.path.join(base_folder, cei_file)

        file_transactions = pd.read_excel(filepath,usecols=cols, engine='openpyxl')
        file_transactions['Ativo'] = file_transactions['Produto'].copy()
        file_transactions['Produto'] = file_transactions['Produto'].apply(lambda t: t.split('-')[0].strip())

        transactions = transactions.append(file_transactions, ignore_index = True)

    transactions.rename(columns = {
        'Entrada/Saída': 'Fluxo',
        'Produto': 'Codigo',
        'Preço unitário': 'Preco',   
        'Valor da Operação': 'Valor Total'
    }, inplace = True)

    transactions['Data'] = pd.to_datetime(transactions['Data'], dayfirst = True)
    transactions.set_index('Data', inplace = True)
    transactions.sort_index(inplace = True)

    for col in ['Fluxo', 'Codigo', 'Ativo']:
        transactions[col] = transactions[col].str.strip()
        
    transactions['Fluxo'].replace({'Credito':'C','Debito':'V'},inplace=True)
    
    transactions["Quantidade"] = transactions["Quantidade"].map(lambda q: q if (type(q) is int) else q.replace(",",".")) # bonificacao tem quantidade com ","
    
    transactions["Quantidade"] = pd.to_numeric(transactions["Quantidade"])
    transactions['Quantidade'] = transactions['Quantidade'] * transactions['Fluxo'].map({"C": 1, "V": -1})
    transactions['Valor Total'] = transactions['Valor Total'] * transactions['Fluxo'].map({"C": 1, "V": -1})
    transactions['Tipo'] = transactions.apply(define_product_type, axis = 1)
    
    

    if save_to_file:
        save_data = transactions.copy()
        save_data.index = save_data.index.date
        save_data.index.name = "Data"
        save_data.reset_index().to_excel('consolidado_cei.xlsx', index = False)
        print('Dados salvos na planilha consolidado_cei.xlsx')

    return transactions


def define_product_type(row):
    is_fii = ("11" in row["Codigo"].upper() or "12" in row["Codigo"].upper()) and ("IMOBILIÁRIO" in row["Ativo"].upper() or "FII" in row["Ativo"].upper())
    movimentacao_is_bonificacao = (row["Movimentação"]=="Bonificação em Ativos")
    movimentacao_is_transf_liquidacao = (row["Movimentação"]=="Transferência - Liquidação")
    
                                                                                  
    if( is_fii and movimentacao_is_transf_liquidacao):
        return "FII"
    if((not is_fii and movimentacao_is_bonificacao) or (not is_fii and movimentacao_is_transf_liquidacao)):
        return "Ação"

    # elif unidecode(row["Mercado"].lower().strip()) in ("opcao de compra", "opcao de venda") and not pd.isna(row["Prazo"]):
    #     return "Opção"
    # elif unidecode(row["Mercado"].lower().strip()) == "exercicio de opcoes":
    #     return "Opção (Exercício)"
    # elif unidecode(row["Mercado"].lower().strip()) in ("mercado a vista", "merc. fracionario"):
    #     return "Ação"
    return "desconhecido"






if __name__ == "__main__":
    consolidate_cei_extracts(save_to_file = True)