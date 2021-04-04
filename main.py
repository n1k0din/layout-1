import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)


excel_table = pandas.read_excel('wine3.xlsx', sheet_name='Лист1')
excel_table.fillna("", inplace=True)
excel_table.rename(
    columns={
        'Категория': 'category',
        'Название': 'title',
        'Сорт': 'sort',
        'Цена': 'price',
        'Картинка': 'image',
        'Акция': 'sale',
    },
    inplace=True
)

wines_from_excel = excel_table.to_dict(orient='record')

wines = defaultdict(list)

for wine in wines_from_excel:
    wines[wine['category']].append(wine)

now = datetime.datetime.now()
founded_at = datetime.datetime(year=1920, month=1, day=1)
delta_years = (now - founded_at).days // 365

template = env.get_template('template.html')

rendered_page = template.render(wines=wines, delta_years=delta_years)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
