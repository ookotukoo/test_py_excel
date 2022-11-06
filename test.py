import requests
import datetime
from bs4 import BeautifulSoup
import pandas as pd

keys_input = {'d1': '',
              'd1mindate': '20091102',
              'd1maxdate': '',
              'd1day': '1',
              'd1month': '',
              'd1year': '',
              'd2': '',
              'd2mindate': '20091102',
              'd2maxdate': '',
              'd2day': '',
              'd2month': '',
              'd2year': '',
              'bSubmit': 'Показать',
              'pge': '1'}

### блок формирования периода даты за прошлый месяц
# получение первой даты
today = datetime.date.today()
first = today.replace(day=1)

keys_input['d1month'] = (first - datetime.timedelta(days=1)).strftime("%m")
keys_input['d1year'] = (first - datetime.timedelta(days=1)).strftime("%Y")

keys_input['d1'] = keys_input['d1year'] + keys_input['d1month'] + '0' + keys_input['d1day']

# получение 2 даты
today = datetime.date.today()
first = today.replace(day=1)

keys_input['d2day'] = (first - datetime.timedelta(days=1)).strftime("%d")
keys_input['d2month'] = (first - datetime.timedelta(days=1)).strftime("%m")
keys_input['d2year'] = (first - datetime.timedelta(days=1)).strftime("%Y")

keys_input['d2'] = keys_input['d2year'] + keys_input['d2month'] + keys_input['d2day']

### блок формирования периода даты за прошлый месяц закончен

datas_for_excel = {'Дата USD/RUB': [], 'Курс USD/RUB': [], 'Время USD/RUB': [], 'Дата JPY/RUB': [], 'Курс JPY/RUB': [], 'Время JPY/RUB': [], 'Результат': []}

# 'Дата JPY/RUB': [], 'Курс JPY/RUB': [], 'Время JPY/RUB': [], 'Результат': []

USD_RUB = {'currency': 'USD_RUB'}

JPY_RUB = {'currency': 'JPY_RUB'}

url = 'https://www.moex.com/ru/derivatives/currency-rate.aspx'


def parsing_USD_RUB():
    response_currency = requests.post(requests.get(url, params=USD_RUB).url, data=keys_input).text

    html_code = BeautifulSoup(response_currency, features='lxml')

    all_table = html_code.find('table', class_='tablels')

    all_line_in_table = all_table.find_all('tr')

    rows_all_table = 1
    cell_count = 0
    rows_count = 0
    give_me_fun = 0

    # разбиваем таблицу на строки
    for line in all_line_in_table:

        # через цикл отбрасываю 2 ненужные строки
        if rows_all_table < 3:
            rows_all_table += 1
        else:
            # разбиваем строки на ячейки
            rows_line_in_table = line.find_all('td')

            for cell_in_rows in rows_line_in_table:

                if cell_count == 0:
                    datas_for_excel['Дата USD/RUB'].append(cell_in_rows.text)
                elif cell_count == 3:
                    datas_for_excel['Курс USD/RUB'].append(float(cell_in_rows.text.replace(',', '.')))
                elif cell_count == 4:
                    datas_for_excel['Время USD/RUB'].append(cell_in_rows.text)
                else:
                    pass
                cell_count += 1
            cell_count = 0
            give_me_fun += 1
        rows_count += 1


def parsing_JPY_RUB():
    response_currency = requests.post(requests.get(url, params=JPY_RUB).url, data=keys_input).text

    html_code = BeautifulSoup(response_currency, features='lxml')

    all_table = html_code.find('table', class_='tablels')

    all_line_in_table = all_table.find_all('tr')

    rows_all_table = 1
    cell_count = 0
    rows_count = 0
    give_me_fun = 0

    # разбиваем таблицу на строки
    for line in all_line_in_table:

        # через цикл отбрасываю 2 ненужные строки
        if rows_all_table < 3:
            rows_all_table += 1
        else:
            # разбиваем строки на ячейки
            rows_line_in_table = line.find_all('td')

            for cell_in_rows in rows_line_in_table:

                if cell_count == 0:
                    datas_for_excel['Дата JPY/RUB'].append(cell_in_rows.text)
                elif cell_count == 3:
                    datas_for_excel['Курс JPY/RUB'].append(float(cell_in_rows.text.replace(',', '.')))
                elif cell_count == 4:
                    datas_for_excel['Время JPY/RUB'].append(cell_in_rows.text)
                    datas_for_excel['Результат'].append(datas_for_excel['Курс USD/RUB'][give_me_fun]/datas_for_excel['Курс JPY/RUB'][give_me_fun])
                else:
                    pass
                cell_count += 1
            cell_count = 0
            give_me_fun += 1
        rows_count += 1


def write_text():

    writer = pd.ExcelWriter('Export_alternative.xlsx', engine='xlsxwriter')

    df = pd.DataFrame(datas_for_excel)

    df.to_excel(writer, sheet_name="Sheet1", index=False)

    worksheet = writer.sheets['Sheet1']

    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
        worksheet.set_column(idx, idx, max_len)

    max_column = df.shape[0] + 1

    for col in df:
        sum_formula = '=SUM(' + col + '2:' + col + str(max_column) + ')'

        df._set_value(str(max_column+1), col, sum_formula)

    writer.close()


if __name__ == '__main__':
    parsing_USD_RUB()

    parsing_JPY_RUB()

    write_text()
