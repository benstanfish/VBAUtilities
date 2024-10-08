from bs4 import BeautifulSoup
import requests

file_name = r'PyProjNetScrape\test\ufc_0.html'

with open(file=file_name) as doc:
    soup = BeautifulSoup(doc, 'html.parser')

for link in soup.find_all('a', href=True, class_='ffclink'):
    # print(link['href'])
    href = link.get('href')
    title = href.split('/')[-1]
    with open(title, 'wb') as f:
        response = requests.get(href)
        f.write(response.content)
