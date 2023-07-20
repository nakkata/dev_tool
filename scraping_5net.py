import requests
import datetime
import csv
import codecs
import time
from bs4 import BeautifulSoup

serch_file = "./serch.txt"


def get_scraping(target_url):
    html = requests.get(str(target_url))
    print(html)
    soup = BeautifulSoup(html.content, "html.parser")
#    elems = soup.find_all(class_="post")
#    elems = soup.select('span')
#    elems = soup.find_all(class_="message")
    elems = soup.find_all(class_=["number","name","date","message"])

    lists = []
    datas = []
    for idx,e in enumerate(elems):
        if((idx > 0) and ((idx % 4) == 0)):
            if(idx > 4000):
                break
            lists.append(datas)
            datas = []
            datas.append(e.getText())
#            print(datas)
        else:
            datas.append(e.getText())
    return lists

def write_data_tsv(fname, datas):
	#開く
#	with codecs.open(fname, 'a', 'CP932','ignore') as f:
	with codecs.open(fname, 'a', 'CP932','ignore') as f:
		writer = csv.writer(f, lineterminator='\n')

		for data in datas:
			writer.writerow(data)
	return 1

if __name__ == '__main__':
    with open(serch_file, 'r', encoding='CP932') as f:
        while True:
            line = f.readline()
            if line:
                datas = []
                print(line)
                datas = get_scraping(line)
                write_data_tsv("./out.csv",datas)
                time.sleep(5)
            else:
                break
