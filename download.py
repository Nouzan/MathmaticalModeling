import shelve
from urllib import request
import time
import os
data = shelve.open('data')
print(len(data['urls']))
count = 0
while count != len(data['urls']):
    count = 0
    for n, u in data['urls']:
        filetype = '.docx'
        if u.endswith('.doc'):
            filetype = '.doc'
        print('Downloading', n + filetype)
        if os.path.exists('download/' + n + filetype):
            print('Has been downloaded')
            count = count + 1
            continue
        req = request.Request(u)
        req.add_header('User-Agent', 'Mozilla/6.0')
        try:
            r = request.urlopen(req).read()
            with open('download/' + n + filetype, 'wb')as file:
                file.write(r)
        except Exception as e:
            print('Failed to download', n)
            pass
#    try:
#        urllib.request.urlretrieve(u, 'download/'+ n +'.docx')
#    except Exception as e:
#        print('Failed to download', n)
#        pass

        time.sleep(2)
    print('Now has been downloaded', count)
data.close()
