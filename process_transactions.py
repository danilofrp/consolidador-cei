import sys, os
import re
import pandas as pd
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from consolidate_cei import consolidate_cei_extracts

def main():
    try:
        action, param = get_args()
    except:
        print(f'\nUsageError: incorrect parameters.\nCorrect usage methods:\npython {__file__} --declaracao [ano]\npython {__file__} --posicao [data:yyyy-mm-dd]')
        return

    transactions = consolidate_cei_extracts(save_to_file = False)

    # transactions = pd.read_excel('consolidado_cei.xlsx', index_col = 'Data', parse_dates = True).sort_index()

    if action == '--declaracao':
        declaration, realised_monthly, realised_monthly_fii = get_declaration_info(transactions, param)

        filename = f'declaracao_{param}.xlsx'
        
        writer = pd.ExcelWriter(filename)
        declaration.to_excel(writer, 'Declaração de Bens')
        realised_monthly.to_excel(writer, 'Lucro Realizado Total')
        realised_monthly_fii.to_excel(writer, 'Lucro Realizado (Somente FIIs)')
        writer.save()

        beutify_positions_excel(filename, history_column = 'F')
        print(f'Declaracao salva em {filename}')

    elif action == '--posicao':
        positions, realised_monthly, realised_monthly_fii = get_position_info(transactions, param)
        
        filename = f'posicoes_{param}.xlsx'
        positions.to_excel(filename, index = False)

        writer = pd.ExcelWriter(filename)
        positions.to_excel(writer, 'Posições')
        realised_monthly.to_excel(writer, 'Lucro Realizado Total')
        realised_monthly_fii.to_excel(writer, 'Lucro Realizado (Somente FIIs)')
        writer.save()

        beutify_positions_excel(filename, history_column = 'F')
        print(f'Posicoes salvas em {filename}')
    
    return

def get_args():
    action = sys.argv[1]
    if action == '--declaracao':
        try:
            year = int(sys.argv[2])
        except:
            year = date.today().year - 1
        return action, year
    elif action == '--posicao':
        try:
            limit_date = sys.argv[2]
        except:
            limit_date = date.today().strftime('%Y-%m-%d')
        return action, limit_date
    else:
        raise Exception()


def get_declaration_info(transactions, interest_year):
    previous_year = interest_year - 1

    positions_previous_year, realised_monthly_previous_year, realised_monthly_fii_previous_year = get_position_info(transactions, f'{previous_year}-12-31')
    positions_interest_year, realised_monthly_interest_year, realised_monthly_fii_interest_year = get_position_info(transactions, f'{interest_year}-12-31', ignore_history_previous_to = interest_year)

    declaration = prepare_declaration_dataframe(positions_previous_year, positions_interest_year, interest_year)
    return declaration, realised_monthly_interest_year, realised_monthly_fii_interest_year


def prepare_declaration_dataframe(positions_previous_year, positions_interest_year, interest_year):
    previous_year = interest_year - 1
    previous_year_declaration = pd.DataFrame()
    previous_year_declaration[f'Posição em {previous_year}-12-31'] = positions_previous_year['Valor Total']

    interest_year_declaration = pd.DataFrame()
    interest_year_declaration[f'Posição em {interest_year}-12-31'] = positions_interest_year['Valor Total']
    interest_year_declaration[f'Quantidade em {interest_year}-12-31'] = positions_interest_year['Quantidade']
    interest_year_declaration[f'Preço médio em {interest_year}-12-31'] = positions_interest_year['Preço Médio']
    interest_year_declaration[f'Histórico em {interest_year}-12-31'] = positions_interest_year['Historico']

    declaration = pd.merge(previous_year_declaration, interest_year_declaration, left_index = True, right_index = True, how = 'outer')
    declaration.sort_index(inplace = True)
    declaration.fillna(0, inplace = True)
    declaration = declaration[(declaration[f'Posição em {previous_year}-12-31'] != 0) | (declaration[f'Posição em {interest_year}-12-31'] != 0)]
    declaration.index.name = 'Ativo'

    return declaration


def get_position_info(transactions, limit_date, ignore_history_previous_to = 1900):
    ignore_history_previous_to = int(ignore_history_previous_to)

    positions = {}
    realised_monthly = pd.Series(name = 'Realizado Mensal', dtype = "float64")
    realised_monthly_fii = pd.Series(name = 'Realizado Mensal', dtype = "float64")

    for index, transaction in transactions[:limit_date].iterrows():
        ignore_history = ignore_history_previous_to > index.year
        
        codigo = transaction['Codigo']
        if codigo == 'USIMA96E':
            codigo = 'USIM5'
        
        if codigo not in positions:
            position = {
                'asset': codigo, 
                'qtd': 0,
                'preco_medio': 0,
                'status': None,
                'historico': []
            }
        else:
            position = positions[codigo]

        if transaction['Fluxo'] == 'C':
            position, realised = process_buy(index, position, transaction, ignore_history)
        elif transaction['Fluxo'] == 'V':
            position, realised = process_sell(index, position, transaction, ignore_history)

        position = update_position_status(position)

        realised_monthly = realised_monthly.append(
                pd.Series([realised], name = "Realizado Mensal", index = [index])
            )
        if "FII" in transaction["Ativo"] and "CI" in transaction["Ativo"]:
            realised_monthly_fii = realised_monthly_fii.append(
                pd.Series([realised], name = "Realizado Mensal", index = [index])
            )
        
        positions[codigo] = position

    realised_monthly = realised_monthly.groupby(pd.Grouper(freq = "M")).sum().resample("M").asfreq().fillna(0)
    realised_monthly.index = realised_monthly.index.date
    realised_monthly.index.name = "Mês"

    realised_monthly_fii = realised_monthly_fii.groupby(pd.Grouper(freq = "M")).sum().resample("M").asfreq().fillna(0)
    realised_monthly_fii.index = realised_monthly_fii.index.date
    realised_monthly_fii.index.name = "Mês"
    positions_df = prepare_position_dataframe(positions)
    return positions_df, realised_monthly.round(2), realised_monthly_fii.round(2)


def process_buy(index, position, transaction, ignore_history = False):
    realised = 0
    if position['status'] != 'short':
        position['preco_medio'] = (position['qtd']*position['preco_medio'] + transaction['Quantidade']*transaction['Preco'])/(position['qtd'] + transaction['Quantidade'])
    else:
        realised = transaction['Quantidade'] * (position['preco_medio'] - transaction['Preco'])
    position['qtd'] = position['qtd'] + transaction['Quantidade']
    if not ignore_history:
        position['historico'].append(f'{index.date()} Compra de {position["asset"]} ({transaction["Quantidade"]} x {transaction["Preco"]:.2f})')

    return position, realised


def process_sell(index, position, transaction, ignore_history = False):
    realised = 0
    if position['status'] == 'long':
        realised = (-transaction['Quantidade']) * (transaction['Preco'] - position['preco_medio'])
    else:
        position['preco_medio'] = ((-position['qtd'])*position['preco_medio'] - transaction['Quantidade']*transaction['Preco'])/((-position['qtd']) - transaction['Quantidade'])
    position['qtd'] = position['qtd'] + transaction['Quantidade']
    if not ignore_history:
        position['historico'].append(f'{index.date()} Venda de {position["asset"]} ({-transaction["Quantidade"]} x {transaction["Preco"]:.2f})')

    return position, realised


def update_position_status(position):
    if position['qtd'] == 0:
        position['preco_medio'] = 0
        position['status'] = None
    elif position['qtd'] > 0:
        position['status'] = 'long'
    elif position['qtd'] < 0:
        position['status'] = 'short'
    return position


def prepare_position_dataframe(positions):
    positions_df = pd.DataFrame(positions).T.sort_index()
    positions_df["preco_medio"] = positions_df["preco_medio"].astype(float).round(2)
    positions_df['Valor Total'] = positions_df['preco_medio'] * positions_df['qtd']
    positions_df.rename(columns = {'asset': 'Ativo', 'preco_medio': 'Preço Médio', 'qtd': 'Quantidade'}, inplace = True)
    positions_df['Historico'] = positions_df['historico'].apply(lambda x: str.join('\n', x))
    positions_df = positions_df[['Ativo', 'status', 'Preço Médio', 'Quantidade', 'Valor Total', 'Historico']]
    return positions_df


def beutify_positions_excel(filename, history_column):
    wb = load_workbook(filename)
    ws = wb.active
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
        ws.column_dimensions[col].width = 20
    ws.column_dimensions[history_column].width = 50
    for cell in ws[history_column]:
        cell.alignment = Alignment(wrapText=True)
    wb.save(filename)



if __name__ == "__main__":
    main()