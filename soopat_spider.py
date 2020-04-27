import re
import time

from excel import generate_excel, read_xlrd
import requests
from lxml import etree


class Soopat_Spider(object):
    @staticmethod
    def run(content):
        html = Soopat_Spider.get_html(content, 1)
        html = etree.HTML(html)
        try:
            number = int(html.xpath('//div[@class="menu"]/p/b[1]/text()')[0])
        except Exception as e:
            print('错误原因：{}'.format(e))
            print('请打开浏览器输入验证码，正在进行{}的爬取'.format(content))
            exit()
        result = []
        if number > 1000:
            number = 1000
        for i in range(0, number, 10):
            html = Soopat_Spider.get_html(content, i)
            html = etree.HTML(html)
            result += Soopat_Spider.analysis_html(html)
            time.sleep(3)
        try:
            if result:
                generate_excel(result, content)
        except Exception as e:
            print('写入文件错误，错误原因：{}'.format(e))

    @staticmethod
    def get_html(content, count):
        url = 'http://www.soopat.com/Home/Result?SearchWord={}&FMZL=Y&SYXX=Y&WGZL=Y&FMSQ=Y&PatentIndex={}'
        headers = {
            'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36'
        }
        proxies = {
            'http': 'http://192.168.118.113:38812',
            'https': 'http://192.168.118.113:38812'
        }
        html = requests.get(url.format(content, count), headers=headers, proxies=proxies).text
        time.sleep(1.5)
        return html

    @staticmethod
    def analysis_html(html):
        divs = html.xpath('//div[@style="min-height: 180px;max-width: 1080px;"]')
        result = []
        for div in divs:
            try:
                title = div.xpath('./div[2]/h2/font/text()')[1] + div.xpath('./div[2]/h2/a/text()')[0]
                url = 'http://www.soopat.com' + div.xpath('./div[2]/h2/a/@href')[0] + \
                      div.xpath('./div[2]/h2/a/font/text()')[0]
                span = div.xpath('string(./div[2]/span[1])')
                applicant = re.findall('申请人：(.*?)- 申请日', span, re.S)
                if applicant:
                    applicant = applicant[0]
                else:
                    applicant = ''
                filing_date = re.findall('申请日：(.*?)- 主分类号', span, re.S)
                if filing_date:
                    filing_date = filing_date[0]
                else:
                    filing_date = ''
                main_classification_number = re.findall('- 主分类号：(.*?)\s.', span, re.S)
                if main_classification_number:
                    main_classification_number = main_classification_number[0]
                else:
                    main_classification_number = ''
            except Exception as e:
                print(e)
            inventor = div.xpath('string(./div[2]/span[2])')
            inventor = re.findall('发明人：(.*?)摘要:', inventor, re.S)
            if inventor:
                inventor = inventor[0]
            else:
                inventor = applicant
            item = {
                'title': title,
                'url': url,
                'filing_date': filing_date,
                'applicant': applicant,
                'main_classification_number': main_classification_number,
                'inventor': inventor,
            }
            print(item)
            result.append(item)
        return result


if __name__ == '__main__':
    for i in read_xlrd(excelFile='./机械运载学部2019版.xlsx'):
        content = i[1]
        Soopat_Spider.run(content)
