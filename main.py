import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
import logging

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s; %(levelname)s; %(name)s; %(message)s',
    filename='logs.lod',
    filemode='w',
    )

this_year = int(datetime.datetime.now().year)
year_of_born = 1920
delta = this_year - year_of_born
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

logging.info(wine_by_category)
with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8001), SimpleHTTPRequestHandler)
server.serve_forever()
