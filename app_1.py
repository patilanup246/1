import requests
from bs4 import BeautifulSoup
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
import re
import pandas as pd
import json
import logging
import sys
import time

# Create a custom logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Create handlers
c_handler = logging.StreamHandler(sys.stdout)
f_handler = logging.FileHandler(f'{logger.name}.log')
c_handler.setLevel(logging.DEBUG)
f_handler.setLevel(logging.DEBUG)

# Create formatters and add it to handlers
c_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
f_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
c_handler.setFormatter(c_format)
f_handler.setFormatter(f_format)

# Add handlers to the logger
logger.addHandler(c_handler)
logger.addHandler(f_handler)

root_dir = os.path.dirname(os.path.abspath(__file__))
def num(s):
    try:
        return int(s.replace(",","").strip())
    except ValueError:
        try:
            return float(s.replace(",","").strip())
        except ValueError:
            return s

def amfiindia():
    try:
        IsAll=False
        IsAllBreak = False

        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Referer": "https://www.amfiindia.com/research-information/aum-data/average-aum",
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
        }
        requests.get("https://www.amfiindia.com/research-information/aum-data/average-aum", timeout=10)
        url = "https://www.amfiindia.com/modules/AverageAUMDetails"
        responsejsons =[{"Selected":False,"Text":"April 2024 - March 2025","Value":"0"},{"Selected":False,"Text":"April 2023 - March 2024","Value":"0"},{"Selected":False,"Text":"April 2022 - March 2023","Value":"0"},{"Selected":False,"Text":"April 2021 - March 2022","Value":"0"},{"Selected":False,"Text":"April 2020 - March 2021","Value":"0"},{"Selected":False,"Text":"April 2019 - March 2020","Value":"1"},{"Selected":False,"Text":"April 2018 - March 2019","Value":"2"},{"Selected":False,"Text":"April 2017 - March 2018","Value":"3"},{"Selected":False,"Text":"April 2016 - March 2017","Value":"4"},{"Selected":False,"Text":"April 2015 - March 2016","Value":"5"},{"Selected":False,"Text":"April 2014 - March 2015","Value":"6"},{"Selected":False,"Text":"April 2013 - March 2014","Value":"7"},{"Selected":False,"Text":"April 2012 - March 2013","Value":"8"},{"Selected":False,"Text":"April 2011 - March 2012","Value":"9"}]
        for item in responsejsons:
            try:
                ValueCat = item['Value']
                TextY=item['Text'].strip().split("-")
                TextM1 = TextY[0].strip().split(" ")
                TextM2 = TextY[1].strip().split(" ")
                ListQ=['January - March '+str(TextM2[1]),'October - December '+str(TextM1[1]),'July - September '+str(TextM1[1]),'April - June '+str(TextM1[1])]
                for itemListQ in ListQ:
                    year_quater = itemListQ
                    xsplit = year_quater.split("-")
                    date_str = '01-' + xsplit[1].strip().replace(" ", "-")
                    filename = datetime.strptime(date_str, '%d-%B-%Y').date()
                    if len(str(filename.month))==1:
                        mnts='0'+str(filename.month)
                    else:
                        mnts = str(filename.month)

                    if len(str(filename.day))==1:
                        days='0'+str(filename.day)
                    else:
                        days = str(filename.day)

                    datestring = str(filename.year)+ '-' +  mnts + '-' + days
                    mainList = []
                    datad = {'AUmType': 'S',
                             'AumCatType': 'Categorywise',
                             'MF_Id': -1,
                             'Year_Id': ValueCat,
                             'Year_Quarter': year_quater}
                    response = requests.post(url=url, data=datad, headers=headers)
                    soup = BeautifulSoup(response.text, "html.parser")
                    hdnVal = BeautifulSoup(soup.find('div', {'id': "divExcel"}).decode_contents().replace("\n", ""),
                                           "html.parser")
                    tags = hdnVal.find_all('tr')
                    for tag in tags:
                        if tag.text == '' or 'Fund Of Funds - Domestic' in tag.text:
                            continue
                        if 'AMFI Code' in tag.text:
                            columns = ['AMFI Code', 'Scheme NAV Name',
                                       'Average AUM for The Month- Excluding Fund of Funds - Domestic but including Fund of Funds - Overseas',
                                       'Average AUM for The Month- Fund Of Funds - Domestic']

                            logger.info(f'columns: {columns}')
                            mainList.append(columns)
                            continue

                        if len(tag.find_all('th')) > 0:
                            if 'Mutual Fund Total' in tag.text or 'Grand Total' in tag.text:
                                columns = []
                                columns.append('')
                                for ta in tag.find_all('th'):
                                    columns.append(num(ta.text))
                                logger.info(f'columns: {columns}')
                                mainList.append(columns)
                                continue

                            columns = []
                            for ta in tag.find_all('th'):
                                columns.append(ta.text)

                            logger.info(f'columns: {columns}')
                            mainList.append(columns)
                            continue

                        if len(tag.find_all('td')) > 0:
                            columns = []
                            for ta in tag.find_all('td'):
                                columns.append(num(ta.text))
                            logger.info(f'columns: {columns}')

                            mainList.append(columns)

                    if ['No records to display'] in mainList:
                        continue;

                    df1 = pd.DataFrame(mainList)
                    df1.to_excel(datestring + '.xlsx', index=False)
                    if IsAll == False:
                        IsAllBreak=True
                        break
                if IsAllBreak == True:
                    break
            except Exception as e:
                logger.error(f'{str(e)}')
                continue



    except Exception as e:
        logger.error(f'{str(e)}')

if __name__ == '__main__':
    logger.info(f'App Start')
    amfiindia()
    # import os, os.path
    # import win32com.client
    #
    # if os.path.exists("F:\\upwork_2020\\truenorthmanagersllp\\amfiindia\\final\\Template.xlsm"):
    #     xl = win32com.client.Dispatch("Excel.Application")
    #     xl.Workbooks.Open(os.path.abspath("F:\\upwork_2020\\truenorthmanagersllp\\amfiindia\\final\\Template.xlsm"), ReadOnly=1)
    #     xl.Application.Run("F:\\upwork_2020\\truenorthmanagersllp\\amfiindia\\final\\Template.xlsm!Module1.amfi_automate")
    #     ##    xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    #     xl.Application.Quit()  # Comment this out if your excel script closes
    #     del xl
    logger.info(f'App End')
