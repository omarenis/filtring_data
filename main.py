# This is a sample Python script.
import pandas as pd
import openpyxl
data = {}
def get_formatted_data():
    book = openpyxl.load_workbook('Extraction-commandes-20221221.xlsx')
    sheet = book[book.sheetnames[0]]
    cols = []
    for cell in list(sheet.rows)[1]:
        data[cell.value] = []
        cols.append(cell.value)

    for row in list(sheet.rows)[2:]:
        for cell, key in zip(row, list(data.keys())):
            data[key].append(cell.value)
    return pd.DataFrame(data), cols


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
def read_commands():
    return pd.read_excel('Extraction-commandes-20221221.xlsx')


def read_list_clients():
    dataframe = pd.read_excel('Liste-clients.xlsx')
    return list(dataframe['Liste Clients'])

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    data, cols = get_formatted_data()
    clients = read_list_clients()
    rslt_df = data[data['SalesOrder.AccountReference'].isin(clients)]
    print(rslt_df)
    rslt_df.to_csv('output.csv')
