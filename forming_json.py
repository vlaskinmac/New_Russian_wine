import collections
import json
from collections import defaultdict, ChainMap
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pprint import pprint

from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
import os.path


def count_year():
    today = datetime.now()
    date = datetime(1920, 1, 1)
    period = (today - date)
    year = period.days//365
    return year


def grouping_catalog_wine(filename):
    excel_data_df = pandas.read_excel(filename, na_values=None, keep_default_na=False)
    excel_data_with_wine = excel_data_df.to_dict('records')
    result = defaultdict(list)
    for wine in excel_data_with_wine:
        category = wine["Категория"]
        result[category].append(wine)
    return result


def rendering_site(filepath):
    year = count_year()
    catalog_wine = grouping_catalog_wine(filename=filepath)
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        age_company_text=f"Уже {year} год с вами",
        data_wine=catalog_wine
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


def main():
    filepath = os.path.abspath('wine2.xlsx')
    rendering_site(filepath)
    server = HTTPServer(('127.0.0.1', 8002), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()