import time

import xlwt
from selenium import webdriver

from excel import read_xlrd


class spider(object):
    def __init__(self):
        # self.url = 'https://kns8.cnki.net/kns/defaultresult/index'
        self.driver = webdriver.Chrome(executable_path='/home/jiuzhang/jobs/blueVipCerifica/chrome_local/chromedriver')

    def data_write(self,file_path, datas):
        f = xlwt.Workbook()
        sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

        # 将数据写入第 i 行，第 j 列
        i = 0
        for data in datas:
            for j in range(len(data)):
                sheet1.write(i, j, data[j])
            i = i + 1

        f.save(file_path)  # 保

    def run(self):
       self.driver.get('https://kns8.cnki.net/kns/defaultresult/index')
       time.sleep(2)
       self.driver.find_element_by_xpath('//span[@value="SU$%=|"]').click()
       time.sleep(1)
       self.driver.find_element_by_xpath('//li[@data-val="TI"]/a').click()
       for i in read_xlrd(excelFile='./郭东明.xlsx'):
           time.sleep(2)
           self.driver.find_element_by_id('txt_search').send_keys(i[0])
           self.driver.find_element_by_class_name('search-btn').click()
           time.sleep(5)
           print(i)
           for x in range(len(self.driver.find_elements_by_xpath('//a[@class="KnowledgeNetLink"]'))-1):
                div = self.driver.find_elements_by_xpath('//a[@class="KnowledgeNetLink"]|//a[@style="cursor:default;color:#778192;"]') #把查找元素的语句移到循环内
                if div[x].text=='郭东明':
                    continue
                try:
                    if  div[x].text[0] in '0123456789':
                        continue
                    div[x].click()
                except:
                    self.driver.find_element_by_xpath('//a[@class="showAllAuthors"]').click()
                    div[x].click()
                try:
                    self.driver.switch_to.window(self.driver.window_handles[1])#切换到点开的页面句柄下进行操作
                except:
                    continue
                print(self.driver.find_element_by_id('showname').text+":"+self.driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[1]/div/div[3]/h3[1]/span/a').text)
                self.driver.close()#关闭当前标签页
                time.sleep(1.5)
                self.driver.switch_to.window(self.driver.window_handles[0])
                time.sleep(1.5)

           self.driver.find_element_by_id('txt_search').clear()
if __name__ == '__main__':
    s=spider()
    s.run()