
# python imports
import xlsxwriter 
import pandas as pd
import openpyxl as op
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver


class google_trends:
    data = []
    source = None
    browser = None
    filename = "data.xlsx"
    headers = ['Item', 'Date', 'Count', 'URL']
    
    def __init__(self):
        self.browser = webdriver.Chrome('chromedriver.exe')
        self.browser.set_window_position(450, 30)
        self.browser.set_window_size(900, 700)

    def start(self):
        url     = 'https://trends.google.com/trends/explore?q=cars&geo=US'
        sheet   = pd.read_excel('queries.xlsx', 'Main')
        items   = sheet['Keyword']
        google  = sheet['Google URL']
        youtube = sheet['YouTube URL']
        self.create_excel_file('Google')
        self.browser.get(url)
        for index in range(3):  
            print(f'\t\t{items[index].upper()}')
            self.get_data('Google', items[index], google[index])

    def get_data(self, search, item, url):
        self.browser.get(url)
        sleep(5)
        soup    = BeautifulSoup(self.browser.page_source, features='lxml')
        table   = soup.find('table')
        try:
            tbody   = table.find('tbody')
        except:
            return
        trs     = tbody.find_all('tr')
        for tr in trs:
            value = {}
            td = tr.find_all('td')
            date = td[0].get_text()
            year = date.split(',')[1].strip(' ')
            year = int(year.replace('\u202c', ''))
            if year > 2003:
                print(date, ' = ', td[1].get_text())
                value['count']  = td[1].get_text()
                value['date']   = date
                value['item']   = item
                value['url']    = url
                self.data.append(value)
                self.write_to_excel(search, value)


    def create_excel_file(self, search):
        # creating new excle file
        workbook = xlsxwriter.Workbook(self.filename)
        if search == 'Google':  workbook.add_worksheet('Google')
        if search == 'YouTube': workbook.add_worksheet('YouTube')
        workbook.close()
        workbook = op.load_workbook(self.filename, False)
        if search == 'Google':  worksheet = workbook['Google']
        if search == 'YouTube': worksheet = workbook['YouTube']
        worksheet.append(self.headers)
        workbook.save(self.filename)
        workbook.close()

    def write_to_excel(self, search, value):
        workbook = op.load_workbook(self.filename, False)
        if search == 'Google':  worksheet = workbook['Google']
        if search == 'YouTube': worksheet = workbook['YouTube']
        items = []
        items.append(value['item'])
        items.append(value['date'])
        items.append(value['count'])
        items.append(value['url'])
        worksheet.append(items)
        workbook.save(self.filename)
        workbook.close()


if __name__ == '__main__':
    gtrend = google_trends()
    gtrend.start()


