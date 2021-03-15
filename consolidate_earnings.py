import os
import re
import xlrd
import pandas as pd

interest_columns_provisionados = ['Ativo','Previsão de Pagamento', 'Valor']
interest_columns_creditados = ['Ativo', 'Creditado No Mês']

def consolidate_cei_earnings(base_folder = "extratos_mensais", save_to_file = True):

    cei_files = os.listdir(base_folder)

    earnings_provisionados = pd.DataFrame()
    earnings_creditados = pd.DataFrame()
    for cei_file in cei_files:
        if (cei_file == '.DS_Store') or (".~" in cei_file): continue # for macOS compatibility and temporary files

        month = re.search(r'(.*)_extrato_cei_.*\.xlsx?$', cei_file)[1]
        broker = re.search(r'extrato_cei_(.*)\.xlsx?$', cei_file)[1]
        filepath = os.path.join(base_folder, cei_file)

        wb = xlrd.open_workbook(filepath, logfile=open(os.devnull, 'w'))
        df = pd.read_excel(wb, engine='xlrd').dropna(how = "all", axis = 1)

        index_provisionados = None
        index_total_provisionados = None
        for index, row in df.iterrows():
            if "PROVENTOS EM DINHEIRO - PROVISIONADOS" in row.values:
                index_provisionados = index
            if "TOTAL PROVISIONADO" in row.values:
                index_total_provisionados = index
        
        if index_provisionados is not None:
            provisionados = df.loc[index_provisionados:]
            provisionados.columns = df.loc[index_provisionados + 1]
            provisionados = provisionados.loc[index_provisionados + 2 : index_total_provisionados - 1, interest_columns_provisionados]
            provisionados["Mês"] = month
            provisionados["Tipo"] = "Provisionado"
            provisionados["Corretora"] = broker

            earnings_provisionados = earnings_provisionados.append(provisionados, ignore_index = True)

        index_creditados = None
        index_total_creditados = None
        for index, row in df.iterrows():
            if "PROVENTOS EM DINHEIRO - CREDITADOS" in row.values:
                index_creditados = index
            if "TOTAL CREDITADO" in row.values:
                index_total_creditados = index
        

        if index_creditados is not None:
            creditados = df.loc[index_creditados:]
            creditados.columns = df.loc[index_creditados + 1]
            creditados = creditados.loc[index_creditados + 2 : index_total_creditados - 1, interest_columns_creditados]
            creditados.rename(columns = {"Creditado No Mês": "Valor"}, inplace = True)
            creditados["Mês"] = month
            creditados["Tipo"] = "Creditado"
            creditados["Corretora"] = broker

            earnings_creditados = earnings_creditados.append(creditados, ignore_index = True)

    filename = "consolidado_proventos.xlsx"
    writer = pd.ExcelWriter(filename)

    if not earnings_creditados.empty:
        earnings_creditados = earnings_creditados[["Mês", "Corretora", "Ativo", "Valor"]]
        earnings_creditados_by_month = earnings_creditados.groupby(["Mês"])[["Valor"]].sum().reset_index()
        earnings_creditados_by_asset = earnings_creditados.groupby(["Ativo", "Corretora"])[["Valor"]].sum().reset_index()

        earnings_creditados.to_excel(writer, 'Creditados', index = False)
        earnings_creditados_by_month.to_excel(writer, 'Creditados Por Mês', index = False)
        earnings_creditados_by_asset.to_excel(writer, 'Creditados Por Ativo', index = False)

    if not earnings_provisionados.empty:
        earnings_provisionados = earnings_provisionados[["Mês", "Corretora", "Ativo", "Previsão de Pagamento", "Valor"]]
        earnings_provisionados_by_month = earnings_provisionados.groupby(["Mês"])[["Valor"]].sum().reset_index()
        earnings_provisionados_by_asset = earnings_provisionados.groupby(["Ativo", "Corretora"])[["Valor"]].sum().reset_index()

        earnings_provisionados.to_excel(writer, 'Provisionados', index = False)
        earnings_provisionados_by_month.to_excel(writer, 'Provisionados Por Mês', index = False)
        earnings_provisionados_by_asset.to_excel(writer, 'Provisionados Por Ativo', index = False)

    writer.save()
    print('Dados salvos na planilha consolidado_proventos.xlsx')

if __name__ == "__main__":
    consolidate_cei_earnings(save_to_file = True)