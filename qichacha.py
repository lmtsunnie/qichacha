import requests
import re
import sys
import xlrd
import xlwt

wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
ws = wb.add_sheet('test', cell_overwrite_ok=True)
ws.write(0, 0, label='企业名称')
ws.write(0, 1, label='法人代表')
ws.write(0, 2, label='大股东')

data = xlrd.open_workbook('/Users/sunnie/Downloads/test.xlsx')
table = data.sheets()[0]
row = table.nrows  # 行数
col = table.ncols  # 列数

def get_para(company):
    url = "https://www.qichacha.com/search"

    querystring = {"key": company}

    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",

        "Accept-Encoding": "gzip, deflate, br",

        "Accept-Language": "zh",

        "Cache-Control": "no-cache",

        "Connection": "keep-alive",

        "Cookie": "acw_tc=6548caa715593109996256001e4b1be43dfef1fed3a7782c2b98cb26be; QCCSESSID=prj5idn7p6htstpepopn13v9c7; UM_distinctid=16b0e2da6188e8-03f15580bab44a-37647e03-1aeaa0-16b0e2da61972a; zg_did=%7B%22did%22%3A%20%2216b0e2da886abe-00e7ffc1e3ebbc-37647e03-1aeaa0-16b0e2da887b17%22%7D; _uab_collina=155931100201874815644592; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1559311003; CNZZDATA1254842228=2138337416-1559309156-%7C1559352377; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201559356886681%2C%22updated%22%3A%201559356893825%2C%22info%22%3A%201559311001742%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.qichacha.com%22%2C%22cuid%22%3A%20%229ca513a45c0b446e920aa4b8a26f5e83%22%7D; hasShow=1; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1559356894",

        "Host": "www.qichacha.com",

        "Pragma": "no-cache",

        "Referer": "https://www.qichacha.com",

        "Upgrade-Insecure-Requests": "1",

        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)
    print(response.text)
    print("--------")
    para = re.findall('href="/firm_(.*?).html" target="_blank" class="ma_h1">', response.text, re.S)
    if not para:
        return
    else:
        print(para[0])
        return para[0]


def get_res1(para):
    urls = "https://www.qichacha.com/firm_{}.html".format(para)

    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",

        "Accept-Encoding": "gzip, deflate, br",

        "Accept-Language": "zh",

        "Cache-Control": "no-cache",

        "Connection": "keep-alive",

        "Cookie": "acw_tc=6548caa715593109996256001e4b1be43dfef1fed3a7782c2b98cb26be; QCCSESSID=prj5idn7p6htstpepopn13v9c7; UM_distinctid=16b0e2da6188e8-03f15580bab44a-37647e03-1aeaa0-16b0e2da61972a; zg_did=%7B%22did%22%3A%20%2216b0e2da886abe-00e7ffc1e3ebbc-37647e03-1aeaa0-16b0e2da887b17%22%7D; _uab_collina=155931100201874815644592; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1559311003; CNZZDATA1254842228=2138337416-1559309156-%7C1559352377; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201559356886681%2C%22updated%22%3A%201559356893825%2C%22info%22%3A%201559311001742%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.qichacha.com%22%2C%22cuid%22%3A%20%229ca513a45c0b446e920aa4b8a26f5e83%22%7D; hasShow=1; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1559356894",

        "Host": "www.qichacha.com",

        "Pragma": "no-cache",

        "Referer": "https://www.qichacha.com",

        "Upgrade-Insecure-Requests": "1",

        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"
    }

    response = requests.request("GET", urls, headers=headers)
    print(response.text)
    print("--------")
    res = re.findall('<h2 class="seo font-20">(.*?)</h2>', response.text, re.S)
    if not res:
        return
    else:
        print(res[0])
        return res[0]


def get_res2(para):
    urls = "https://www.qichacha.com/firm_{}.html".format(para)

    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",

        "Accept-Encoding": "gzip, deflate, br",

        "Accept-Language": "zh",

        "Cache-Control": "no-cache",

        "Connection": "keep-alive",

        "Cookie": "acw_tc=6548caa715593109996256001e4b1be43dfef1fed3a7782c2b98cb26be; QCCSESSID=prj5idn7p6htstpepopn13v9c7; UM_distinctid=16b0e2da6188e8-03f15580bab44a-37647e03-1aeaa0-16b0e2da61972a; zg_did=%7B%22did%22%3A%20%2216b0e2da886abe-00e7ffc1e3ebbc-37647e03-1aeaa0-16b0e2da887b17%22%7D; _uab_collina=155931100201874815644592; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1559311003; CNZZDATA1254842228=2138337416-1559309156-%7C1559352377; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201559356886681%2C%22updated%22%3A%201559356893825%2C%22info%22%3A%201559311001742%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.qichacha.com%22%2C%22cuid%22%3A%20%229ca513a45c0b446e920aa4b8a26f5e83%22%7D; hasShow=1; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1559356894",

        "Host": "www.qichacha.com",

        "Pragma": "no-cache",

        "Referer": "https://www.qichacha.com",

        "Upgrade-Insecure-Requests": "1",

        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"
    }

    response = requests.request("GET", urls, headers=headers)
    print(response.text)
    print("--------")
    res = re.findall(
        '<h3 class="seo font-14">(.*?)</h3></a> <div class="m-t-xs"> <span class="ntag sm text-danger m-r-xs" style="margin-bottom: 2px;">大股东</span>',
        response.text, re.S)
    if not res:
        return
    else:
        print(res[0])
        return res[0]


if __name__ == '__main__':
    # 1到行数-1
    for i in range(1, row):
        company = table.cell(i, 0).value
        para = get_para(company)
        res1 = get_res1(para)
        res2 = get_res2(para)
        ws.write(i, 0, company)
        ws.write(i, 1, res1)
        ws.write(i, 2, res2)

    wb.save('/Users/sunnie/Downloads/res.xls')