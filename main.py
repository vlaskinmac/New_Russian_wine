import os.path
from datetime import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def group_wine_catalog(filename):
    excel_data_df = pandas.read_excel(filename, na_values=None, keep_default_na=False)
    group_of_wines = excel_data_df.to_dict("records")
    grouped_wines = defaultdict(list)
    for wine in group_of_wines:
        category = wine["Категория"]
        grouped_wines[category].append(wine)
    return grouped_wines


def forms_template():
    env = Environment(
        loader=FileSystemLoader("."),
        autoescape=select_autoescape(["html", "xml"])
    )
    return env.get_template("template.html")


def render_site(filepath):
    template = forms_template()
    year_of_opening = 1920
    current_year = datetime.now().year
    year = current_year - year_of_opening
    wine_catalog = group_wine_catalog(filename=filepath)
    wine_catalog_sort = sorted(wine_catalog.items())
    rendered_page = template.render(
        age_of_the_company=f"Уже {year} год с вами",
        wines=dict(wine_catalog_sort).values()
    )
    with open("index.html", "w", encoding="utf8") as file:
        file.write(rendered_page)


def main():
    filepath = os.path.abspath("wines.xlsx")
    render_site(filepath)
    server = HTTPServer(("127.0.0.1", 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
