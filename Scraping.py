# %%
#IMPORT LIBRARIES

from asyncore import write
from pickle import TRUE
from tracemalloc import start
from unittest import skip
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys   
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import numpy as np

import warnings
warnings.filterwarnings("ignore")

# %%
# CONFIGURATION

faculties = ['คณะแพทยศาสตร์', 'คณะทันตแพทยศาสตร์', 'คณะวิศวกรรมศาสตร์']
semester = [1, 2]

# %%
# DO SUM TING WHO CARES
datas = []

PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

for selectedFaculty in faculties:
    for selectedSemester in semester:

        driver.get("https://web.reg.tu.ac.th/registrar/class_info.asp?lang=th")
        Select(driver.find_elements_by_tag_name('select')[0]).select_by_visible_text(selectedFaculty)
        Select(driver.find_elements_by_tag_name('select')[1]).select_by_visible_text(str(selectedSemester))
        Select(driver.find_elements_by_tag_name('select')[2]).select_by_visible_text('2564')
        Select(driver.find_elements_by_tag_name('select')[4]).select_by_visible_text('ทั้งหมด')
        driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]/table/tbody/tr/td/font[3]/input').click()

        pageCount = 0

        data = pd.DataFrame(
            {
                'campas': [],
                'courseYear': [],
                'faculty': [],
                'semester': [],
                'quota': [],
                'classCode': [],
                'className': [],
                'instructorName': [],
                'credit': [],
                'section': [],
                # 'classDate': [],
                # 'classRoom': [],
                # 'total': [],
                # 'remaining': [],
                # 'status': [],
            }
        )

        while True:

            outerDataRow = driver.find_elements_by_xpath('/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr')

            for trInOuterDataRow in outerDataRow[3:-1:]:
                tdColumns = trInOuterDataRow.find_elements_by_tag_name('td')
                subjectDetails = []
                
                subjectDetails.append(tdColumns[1].text) #campus
                subjectDetails.append(tdColumns[2].text) #courseYear
                subjectDetails.append(driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td[2]/div[1]/font/b').text) #faculty
                
                try:
                    if len(tdColumns[3].find_elements_by_tag_name('a')) != 0: #quota
                        subjectDetails.append('Quota')
                    else:
                        subjectDetails.append('None')
                except:
                    subjectDetails.append('None')
                
                subjectDetails.append(tdColumns[4].text) #subjectCode

                subjectDetails.append(tdColumns[5].find_elements_by_tag_name('font')[0].text.splitlines()[0]) #subjectName
                instructorName = ""
                try:
                    if '***' in tdColumns[5].find_elements_by_tag_name('font')[0].text.splitlines()[1]: #instructorName

                        for instructorx in tdColumns[5].find_elements_by_tag_name('font')[0].text.splitlines()[2:]:
                            instructorName += instructorx
                            if instructorx != tdColumns[5].find_elements_by_tag_name('font')[0].text.splitlines()[-1]:
                                instructorName += '\n'
                    else:
                        for instructorx in tdColumns[5].find_elements_by_tag_name('font')[0].text.splitlines()[1:]:
                            instructorName += instructorx
                            if instructorx != tdColumns[5].find_elements_by_tag_name('font')[0].text.splitlines()[-1]:
                                instructorName += '\n'
                except:
                    instructorName = "-"

                subjectDetails.append(instructorName)
                subjectDetails.append(tdColumns[6].text.split()[0]) #credit
                
                subjectDetails.append(tdColumns[7].find_elements_by_tag_name('font')[0].text.splitlines()[0]) #section

                classRoom = ""
                
                try:
                    if tdColumns[8].find_elements_by_tag_name('font')[0].text == "": #classDate
                        subjectDetails.append('N/A')
                        classRoom = 'N/A'
                    else:
                        dates = ""
                        for date in tdColumns[8].find_elements_by_tag_name('font')[0].text.splitlines():
                            # print(date.text.split()+'\n')
                            # dates = date
                            dates += date.split()[0] + ' ' + date.split()[1]

                            if date == tdColumns[8].find_elements_by_tag_name('font')[0].text.splitlines()[0]:
                                classRoom += date.split()[2]                            
                            if date != tdColumns[8].find_elements_by_tag_name('font')[0].text.splitlines()[-1]:
                                dates += '\n'
                        subjectDetails.append(dates)
                except:
                    subjectDetails.append('N/A')
                    classRoom = 'N/AX'
                
                # subjectDetails.append(classRoom)
                # subjectDetails.append(tdColumns[10].text.split()[0])                
                # subjectDetails.append(tdColumns[11].text.split()[0])
                # subjectDetails.append(tdColumns[12].text.split()[0])
                               
                data = data.append({
                    'campas': subjectDetails[0],
                    'courseYear': subjectDetails[1],
                    'faculty': subjectDetails[2],
                    'semester': selectedSemester,
                    'quota': subjectDetails[3],
                    'classCode': subjectDetails[4],
                    'className': subjectDetails[5],
                    'instructorName': subjectDetails[6],
                    'credit': subjectDetails[7],
                    'section': subjectDetails[8],
                    'classDate': subjectDetails[9],
                    # 'classRoom': subjectDetails[10],
                    # 'total': subjectDetails[11],
                    # 'remaining': subjectDetails[12],
                    # 'status': subjectDetails[13]


                }, ignore_index=True)

            time.sleep(0.4)
            pageCount +=1
            clickNext = driver.find_elements_by_xpath("//td[2]/font/a")
            if len(clickNext) >1:
                
                clickNext[1].click()
                continue
            if clickNext[0].text == '[หน้าก่อน]':
                # data.to_excel('data.xlsx', engine='xlsxwriter')
                # sheetName = selectedFaculty + '_' + str(selectedSemester)
                sheetName = selectedFaculty
                # writer = pd.ExcelWriter('data.xlsx')
                datas.append([data, sheetName])
                # data.to_excel(writer, sheet_name=sheetName)
                driver.find_element_by_xpath('//a[contains(text(), "ถอยกลับ")]').click()
                time.sleep(1)
                break;

            clickNext[0].click()

        # print(pageCount)

# dataMerge = []

# for data in datas[::2]:
#     dataMerge.append(pd.concat([data[0], (datas[datas.index(data)+1])[0]]))

writer = pd.ExcelWriter('data1.xlsx', engine='openpyxl')
pd.DataFrame().to_excel(writer)

# for datax in datas[:-1:2]:

#     nextElementIndex = 0
#     for listIndex in range(len(datas)):
#         if str(datas[listIndex]) == str(datax):
#             nextElementIndex = listIndex + 1

#     merge = pd.concat([datax[0], (datas[nextElementIndex])[0]], ignore_index=True, sort=False)
#     merge.index = np.arange(1, len(merge)+1)
#     merge.to_excel(writer, sheet_name=datax[1])
#     datax[0].to_excel(writer, sheet_name=datax[1])

# %%
dataSemester_1 = pd.DataFrame()
dataSemester_2 = pd.DataFrame()

for datax in datas[::2]:
    dataSemester_1 = pd.concat([dataSemester_1, datax[0]], ignore_index=True, sort=False)

for datax in datas[1::2]:
    dataSemester_2 = pd.concat([dataSemester_2, datax[0]], ignore_index=True, sort=False)

dataSemesters = pd.concat([dataSemester_1, dataSemester_2], ignore_index=True, sort=False)
dataSemesters.to_excel(writer, sheet_name='Fact')

writer.save()

print('Finished')

driver.quit()

# %%