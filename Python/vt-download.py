#!/usr/bin/python3

import requests
import time

base_url='https://www.virustotal.com/api/v3/files/%s/download'
apikey=''
hash_list='hashes.txt'
save_dir='samples/VT_%s'

with open(hash_list, 'r') as fp:
    for line in fp:
        file=requests.get(base_url % (line.rstrip()), headers={'x-apikey':apikey})
        open(save_dir % (line.rstrip()), 'wb').write(file.content)
        time.sleep(2)
