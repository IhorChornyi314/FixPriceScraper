import urllib.error
from bs4 import BeautifulSoup as bs
import pandas as pd
from io import BytesIO
from urllib import request
from datetime import date
from PIL import Image


def clean(text):
    return text.replace('\t', '').replace('\n', '').replace('\r', '').strip()


class Scraper:
    def __init__(self):
        self.category_name = ''
        self.subcategory_name = ''
        self.db = pd.DataFrame()

    def collect_subcategories_links(self):
        page = request.urlopen('https://fix-price.ru/catalog/')
        soup = bs(page, 'html.parser')
        subcategory_links = []
        for link in soup.find_all('li', class_='catalog-sub__item'):
            subcategory_links.append('https://fix-price.ru' + link.find('a')['href'])
        return subcategory_links

    def process_page(self, url):
        page = request.urlopen(url)
        soup = bs(page, 'html.parser')
        catalog_box = soup.find('div', id='catalog_sect_cont')
        goods_images = [img.find('img')['src'] for img in catalog_box.find_all('div', class_='product-card-top-container')]
        goods_names = [clean(name.text) for name in catalog_box.find_all('a', class_='product-card__title')]
        goods_prices = [
            clean(f'{price.find_all("span")[0].text} {price.find_all("span")[-1].text}')
            for price in catalog_box.find_all('div', class_='product-card__bottom-badge-price')
        ]
        temp = pd.DataFrame()
        temp['Image'] = pd.Series(goods_images)
        temp['Name'] = pd.Series(goods_names)
        temp['Price'] = pd.Series(goods_prices)
        temp['Category'] = self.category_name
        temp['Subcategory'] = self.subcategory_name
        self.db = self.db.append(temp)

    def process_subcategory(self, url):
        page = request.urlopen(url)
        soup = bs(page, 'html.parser')
        self.category_name = clean(soup.find_all('div', class_='breadcrumb__item')[-2].text)
        self.subcategory_name = clean(soup.find_all('div', class_='breadcrumb__item')[-1].text)
        try:
            number_of_pages = int(soup.find('ul', class_='paging__list').find_all('li')[-1].string)
        except Exception:
            number_of_pages = 1
        for page_number in range(number_of_pages):
            print(f'Processing page number {page_number + 1} out of {number_of_pages}...')
            self.process_page(f'{url.split("?")[0]}?PAGEN_1={str(page_number + 1)}')

    def scrape(self, database_name):
        subcategory_links = self.collect_subcategories_links()
        _ = 0
        for subcategory_link in subcategory_links:
            print(f'Processing subcategory number {_ + 1} out of {len(subcategory_links)}: {subcategory_link}...')
            self.process_subcategory(subcategory_link)
            _ += 1
        self.db.to_csv(database_name, encoding='utf8')


def download_images(database_name):
    df = pd.read_csv(database_name)
    del df['Unnamed: 0']
    writer = pd.ExcelWriter(database_name.replace('.csv', '.xlsx'))
    df.to_excel(writer, 'Sheet1')
    wb = writer.book
    ws = wb.get_worksheet_by_name('Sheet1')
    for _, image_url in zip(range(df['Image'].size), df['Image']):
        print(_ + 1)
        try:
            try:
                image_data = BytesIO(request.urlopen(image_url).read())
            except urllib.error.HTTPError:
                image_data = BytesIO()
                temp = Image.open(BytesIO(request.urlopen(image_url.replace('.JPG', '.webp')).read()))
                temp.save(image_data, format='JPEG')
            img = Image.open(image_data)
            scale = 183 / max(img.size)
            offset_x = (185 - img.size[0] * scale) / 2
            offset_y = (185 - img.size[1] * scale) / 2
            ws.set_row(_ + 1, 138)
            ws.insert_image(_ + 1, 1, image_url, {'image_data': image_data, 'x_scale': scale, 'y_scale': scale,
                                                  'y_offset': offset_y, 'x_offset': offset_x})
        except Exception as e:
            print(e)
    ws.set_row(0, 20)
    ws.set_column(1, 5, 25.5)
    writer.save()
    wb.close()


s = Scraper()
s.scrape('Database%s.csv' % date.today().strftime('%d%m%Y'))
download_images('Database%s.csv' % date.today().strftime('%d%m%Y'))


