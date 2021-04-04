import datetime
from collections import defaultdict
import argparse
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def create_parser():
    parser = argparse.ArgumentParser(description='Запускает сайт про вина')
    parser.add_argument(
        '-i',
        '--input',
        help='Входная excel таблица с данными, wine.xlsx по-умолчанию',
        default='wine.xlsx',
    )

    parser.add_argument(
        '--host',
        help='Интерфейс для запуска сервера, на всех по-умолчанию',
        default='0.0.0.0',
    )

    parser.add_argument(
        '-p',
        '--port',
        help='Порт для запуска сервера, 8000 по-умолчанию',
        type=int,
        default=8000,
    )

    return parser


def get_args_from_parser(parser):
    args = parser.parse_args()

    return args.input, args.host, args.port


def get_winery_age(founded_year=1920):
    return datetime.datetime.now().year - founded_year


if __name__ == '__main__':
    parser = create_parser()
    input_filename, host, port = get_args_from_parser(parser)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    excel_table = pandas.read_excel(input_filename)
    excel_table.fillna("", inplace=True)

    wines_from_excel = excel_table.to_dict(orient='record')

    wines = defaultdict(list)

    for wine in wines_from_excel:
        wines[wine['Категория']].append(wine)

    company_age = get_winery_age()

    template = env.get_template('template.html')

    rendered_page = template.render(wines=wines, delta_years=company_age)

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer((host, port), SimpleHTTPRequestHandler)
    server.serve_forever()
