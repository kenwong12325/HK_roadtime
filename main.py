import requests as req
from bs4 import BeautifulSoup

url = 'https://resource.data.one.gov.hk/td/jss/Journeytimev2.xml'

r = req.get(url)
if r.status_code == 200:
    print("access successfuly")
    with open('traffic.xml', 'w') as f:
        f.write(r.text)
        exec(open('xml2excel.py').read())

