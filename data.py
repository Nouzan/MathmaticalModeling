from urllib import request
from bs4 import BeautifulSoup
import time
import shelve

weburl = 'http://szmqs.gov.cn/xxgk/qt/ztlm/syncpzl/ncp/'
urls = []
docxs = []

dataFile = shelve.open('data')

for i in range(6):
    if i != 0:
        req = request.Request(weburl + 'index_' + str(i) + '.htm')
    else:
        req = request.Request(weburl)
    print('Entering', str(i))
    req.add_header('User-Agent', 'Mozilla/6.0')
    data = request.urlopen(req).read()

    soup = BeautifulSoup(data, "html.parser")


    for url in soup.select('li[class="pclist"] a'):
        if '监测结果' in url.get('title', []):
            urls.append(weburl+url.get('href')[2:])
    # print(urls)
    
    for url in urls:
        reqq = request.Request(url)
        reqq.add_header('User-Agent', 'Mozilla/6.0')
        dat = request.urlopen(reqq).read()
        soupp = BeautifulSoup(dat, "html.parser")
        for uurl in soupp.select('div a'):
            if '.docx' in uurl.getText() or '.doc' in uurl.getText() or '.doc' in uurl.get('href'):
                docxs.append((uurl.getText().split('.')[0], '/'.join(url.split('/')[0:-1]) + uurl.get('href')[1:]))
                print((uurl.getText().split('.')[0], '/'.join(url.split('/')[0:-1]) + uurl.get('href')[1:]))
    time.sleep(2)
print(docxs)
dataFile['urls'] = docxs
dataFile.close()

