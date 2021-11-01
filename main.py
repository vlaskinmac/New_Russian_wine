import os.path
from datetime import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def count_year():
    today = datetime.now()
    date = datetime(1920, 1, 1)
    period = (today - date)
    year = period.days // 365
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
    catalog_wine_sort = sorted(catalog_wine.items())
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        age_company_text=f"Уже {year} год с вами",
        data_wine=dict(catalog_wine_sort).values()
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


def main():
    filepath = os.path.abspath('wine3.xlsx')
    rendering_site(filepath)
    server = HTTPServer(('127.0.0.1', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
