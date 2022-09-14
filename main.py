import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import aspose.words as aw


def check_path(outname):
    outdir = './dir'
    if not os.path.exists(outdir):
        os.mkdir(outdir)
    return os.path.join(outdir, outname)


def save_table(fullname):
    doc = aw.Document(fullname)
    extractedPage = doc.extract_pages(0, 1)
    extractedPage.save(os.path.splitext(fullname)[0] + ".png")


url = 'https://myfin.by/bank/kursy_valjut_nbrb'
response = requests.get(url)
soup = BeautifulSoup(response.text, 'lxml')
rows = soup.find('table', class_='default-table').find_all('tr')
data = {'Валюта': [], 'Курс': [], 'Курс на завтра': [], 'Код': [], 'Единиц': []}
for i in range(1, len(rows)):
    data['Валюта'].append(rows[i].find('a').text)
    data['Курс'].append(rows[i].find_all('td')[1].text)
    data['Курс на завтра'].append(rows[i].find_all('td')[2].text)
    data['Код'].append(rows[i].find_all('td')[3].text)
    data['Единиц'].append(rows[i].find_all('td')[4].text)
df = pd.DataFrame(data, index=[i for i in range(1, len(data['Валюта']) + 1)])
outname = 'exchange_rates.html'

fullname = check_path(outname)
df.to_html(fullname)

save_table(fullname)