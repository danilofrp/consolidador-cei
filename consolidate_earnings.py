import os, sys
import re
import xlrd
import pandas as pd
import numpy as np
from datetime import date


earning_types = ['Dividendo','Juros Sobre Capital Próprio','Rendimento']



def consolidate_cei_earnings(base_folder = "movimentacoes", save_to_file = True):
    
    cols = ['Entrada/Saída','Data','Movimentação','Produto','Instituição','Quantidade','Preço unitário','Valor da Operação']
    

    cei_files = os.listdir(base_folder)

    cei_files = os.listdir(base_folder)
    
    cei_files = [f for f in cei_files if "~$" not in f]
    
    earnings_transactions = pd.DataFrame()
    
    for cei_file in cei_files:
        if (cei_file == '.DS_Store') or ("~lock" in cei_file): continue # for macOS compatibility and removal of temporary files

        filepath = os.path.join(base_folder, cei_file)

        file_transactions = pd.read_excel(filepath,usecols=cols, engine='openpyxl')
        
        file_transactions = file_transactions[file_transactions['Movimentação'].isin(earning_types)]
        
        file_transactions['Ativo'] = file_transactions['Produto'].copy()
        file_transactions['Produto'] = file_transactions['Produto'].apply(lambda t: t.split('-')[0].strip())
        
        earnings_transactions = earnings_transactions.append(file_transactions, ignore_index = True)
        
    earnings_transactions['Data'] = pd.to_datetime(earnings_transactions['Data'], dayfirst = True)
    earnings_transactions.set_index('Data', inplace = True)
    earnings_transactions.sort_index(inplace = True)
    
    

    filename = f"consolidado_proventos.xlsx"
    writer = pd.ExcelWriter(filename)
    
    transaction_years = earnings_transactions.index.year.unique()
    
    for _type in earning_types:
        
        df_consolidated = []
                
        for year in transaction_years:
            
            df_by_year = earnings_transactions[earnings_transactions.index.year==year]
            
            df_by_year_by_type = df_by_year[df_by_year['Movimentação']==_type].groupby(['Produto'])[["Valor da Operação"]].sum().reset_index()
            
            df_by_year_by_type.rename(columns={'Valor da Operação':f'Total Recebido {year}'},inplace=True)
            
            df_by_year_by_type.set_index('Produto',inplace = True)
            
            df_consolidated.append(df_by_year_by_type)
            
        df_consolidated = pd.concat(df_consolidated,axis=1)
        
        df_consolidated['Total'] = df_consolidated.sum(axis=1)
        
        df_consolidated.fillna(0,inplace=True)
        
        
        df_consolidated.to_excel(writer, _type, index = True)
        
    




    writer.save()
    print(f'Dados salvos na {filename}')


    
# sourcery skip: raise-specific-error
if __name__ == "__main__":
    
    consolidate_cei_earnings(save_to_file = True)