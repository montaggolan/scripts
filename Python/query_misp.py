import json
import pymisp
import argparse
import sys
from keys import misp_key, misp_url, misp_verifycert

def importCuckooReport(sourceFile):
    with open(sourceFile) as fp:
        jsonDict=json.load(fp)
    return jsonDict

def queryhashes(misp, hash, report=True, sourceFile=None):
    hashesDict=dict()
    kwargs={"not_tags":"Cuckoo"}
    if report:
        if not sourceFile:
            print("Report file must be given with report option.")
            sys.exit(1)
        hashesDict[hash]=dict()
        src=importCuckooReport(sourceFile)
        for r in src:
            h=r["hashes"][hash]
            response=misp.search_all(h)
            i=0
            for resp in response['response']:
                hashesDict[h+"_"+str(i)]={resp['Event']['id']:[resp['Event']['Org']['name'], resp['Event']['Orgc']['name'],
                resp['Event']['date'], 
                resp['Event']['timestamp'], resp['Event']['publish_timestamp'], resp['Event']['info'], h]}
                i+=1
    return hashesDict

def normalize(i):
    queryables=[]
    if i[0]=="#":
        return queryables
    splitted=i.split('/')
    if "http" in i[:4]:
        if len(splitted)<2:
            return queryables
        queryables.append(splitted[0]+"//"+splitted[2])
        queryables.append(splitted[2])
        indxs=[2,3]
    else:
        queryables.append(splitted[0])
        indxs=[0,1]
    if splitted[-1]!=splitted[indxs[0]] and len(splitted)>indxs[1] and splitted[-1]:
        if "." in splitted[-1]:
            newSplit=splitted[-1].split("?")
            try:
                if newSplit[0][0].lower() in 'abcdefghijklmnoprstuvwxyz123456890':
                    if newSplit[0] not in exceptionList:
                        queryables.append(newSplit[0])
            except:
                pass
    return queryables

def queryUrl(misp, sourceFile):
    hashesDict=dict()
    with open(sourceFile) as fp:
        l=fp.read().splitlines()
    for i in l:
        for j in normalize(i):
            response=misp.search_all(j)
            if response['response']:
                for resp in response['response']:
                    hashesDict[j]={resp['Event']['id']:[resp['Event']['Org']['name'], resp['Event']['Orgc']['name'],
                    resp['Event']['date'], resp['Event']['timestamp'], resp['Event']['publish_timestamp'], resp['Event']['info'], j]}
    return hashesDict

if __name__ == '__main__':
    misp=pymisp.PyMISP(misp_url,misp_key,misp_verifycert,'json')
    #d=queryUrl(misp, sourceFile='submittedUrl_list.txt')
    #with open('result.json', 'w') as fp:
    #    json.dump(d, fp)
    #d=queryhashes(misp, "md5", sourceFile='dataFull.json')
    #with open('result_md5.json', 'w') as fp:
    #    json.dump(d, fp)
    d=queryhashes(misp, "sha1", sourceFile='dataFull.json')
    with open('result_sha1.json', 'w') as fp:
        json.dump(d, fp)
    d=queryhashes(misp, "sha256", sourceFile='dataFull.json')
    with open('result_sha256.json', 'w') as fp:
        json.dump(d, fp)
