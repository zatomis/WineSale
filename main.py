from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import json
import collections
import argparse

START_YEAR = 1920

def createParser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-e', '--excel', type=str, help='Имя файла excel', default='test.xlsx')

    return parser

def correct_year(number):
    years_case = {0:'лет',1:'год',2:'года',3:'года',4:'года',5:'лет',6:'лет',7:'лет',8:'лет',9:'лет'}
    if ((number % 100) >= 11 and (number % 100)<=20):
        return "лет"
    else:
        return years_case.get(number%10)



def main():
    parser = createParser()
    excel_file = parser.parse_args().excel

    env = Environment(loader=FileSystemLoader('.'), autoescape=select_autoescape(['html', 'xml']))
    template = env.get_template('template.html')
    excel_value = pandas.read_excel(excel_file, sheet_name='Лист1',
                                    usecols=['Категория', 'Название', 'Сорт', 'Цена', 'Картинка', 'Акция'])
    json_excel = excel_value.to_json(orient='records', force_ascii=False)

    products_by_category = collections.defaultdict(list)
    for item in json.loads(json_excel):
        category_name = item['Категория']
        product = {
            'Название': item['Название'],
            'Сорт': item['Сорт'],
            'Цена': item['Цена'],
            'Картинка': item['Картинка'],
            'Акция': item['Акция'],
        }

        if category_name not in products_by_category:
            products_by_category[category_name] = []

        products_by_category[category_name].append(product)

    rendered_page = template.render(wines=products_by_category,
                                    count_year=f"{datetime.date.today().year - START_YEAR}  {correct_year(datetime.date.today().year - START_YEAR)}")

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('127.0.0.1', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

if __name__ == '__main__':
    main()


