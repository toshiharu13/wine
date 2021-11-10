import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pprint import pprint

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

today = datetime.datetime.now().year
day_of_born = datetime.datetime(year=1920, month=1, day=1, hour=00).year
delta = today-day_of_born
excel_data_sale = pandas.read_excel(
    'wine3.xlsx', sheet_name='Лист1'
).to_dict(orient='record')
wine_by_category = collections.defaultdict(list)

for wine in excel_data_sale:
    for element in wine:
        if pandas.isna(wine[element]):
            wine[element] = None
    wine_by_category[wine['Категория']].append(wine)

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)
template = env.get_template('template.html')
rendered_page = template.render(
    winery_age=delta,
    excel_data=excel_data_sale,
    wine_by_category=wine_by_category,
)

pprint(wine_by_category)
with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8001), SimpleHTTPRequestHandler)
server.serve_forever()
