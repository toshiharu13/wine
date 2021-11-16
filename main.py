import argparse
import collections
import datetime
import logging
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s; %(levelname)s; %(name)s; %(message)s',
        filename='logs.lod',
        filemode='w',
        )

    parser = argparse.ArgumentParser(
        description='Страничка виноделов Новое русское вино'
    )
    parser.add_argument(
        '--file', help='имя файла xlsx', default='sample_file.xlsx'
    )
    parser.add_argument(
        '--list_in_file', help='имя листа в xlsx файле', default='Лист1'
    )
    args = parser.parse_args()

    this_year = datetime.datetime.now().year
    year_of_born = 1920
    winery_age = this_year - year_of_born
    list_of_wines = pandas.read_excel(
        args.file, sheet_name=args.list_in_file, na_values=None
    ).to_dict(orient='record')
    wine_by_category = collections.defaultdict(list)

    logging.info(list_of_wines)
    for wine in list_of_wines:
        wine_by_category[wine['Категория']].append(wine)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        winery_age=winery_age,
        wine_by_category=wine_by_category.items(),
    )

    logging.info(wine_by_category)
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8001), SimpleHTTPRequestHandler)
    server.serve_forever()
