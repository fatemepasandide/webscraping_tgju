# -*- coding: utf-8 -*-
"""
Created on Sat Aug 15 09:30:52 2020

@author: Fatemeh Pasandideh
"""

import requests
from bs4 import BeautifulSoup
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell



data=[]
week_color=[]
day_color=[]


URL_list=['https://www.tgju.org/profile/ons','https://www.tgju.org/profile/sekee','https://www.tgju.org/profile/geram18','https://www.tgju.org/profile/turkey_usd']


for i in URL_list:
    page = requests.get(i)
    soup = BeautifulSoup(page.text ,'html.parser')
    results = soup.find(id='main')
    job_elems = results.find_all(class_='title')
    
    box_roozane = results.find('div', class_='tgju-widgets-block col-md-12 col-lg-4 tgju-widgets-block-bottom-unset overview-first-block')#Weekly performance
    for elem2 in box_roozane: # finding spans
        td_box = box_roozane.find_all('td')
    
    for i in range(0,18):# choosing 2 first elements
        if i==1 or i==17:
            temp = td_box[i].text
            data.append(temp)
        
    for elem2 in box_roozane: #Find class related to the percentage of last two days
        span_box_from_td_box = box_roozane.find_all('span')
    
    day_color.append(span_box_from_td_box[5].get('class')) #low or high or or without..
    
            
    box_amalkard = results.find('div', class_='tgju-widgets-block col-12 col-md-12 col-lg-6 profile-performance-box')#amalkard hafte
    for elem1 in box_amalkard: # finding spans
        span_box = box_amalkard.find_all('span')
    
    for i in range(0,2):# choosing 2 first elements
        temp = span_box[i].text
        data.append(temp)
        week_color.append(span_box[i].get('class')) # be low or high class to choose color




sep_data=[]
temp=[]
for idx, val in enumerate (data):
    idx=1+idx
    while True:
        temp.append(val)
        if ((idx)%4) == 0:
            sep_data.append(temp)
            temp=[]
            break
        else:
            break



if week_color[0]==['low']:
    gheymat_taghir_ons_hafte = 'red'
elif week_color[0]==['high']:
    gheymat_taghir_ons_hafte = 'green'
else:
    gheymat_taghir_ons_hafte = 'black'
    
if week_color[2]==['low']:
    gheymat_taghir_sekke_hafte = 'red'
elif week_color[2]==['high']:
    gheymat_taghir_sekke_hafte = 'green'
else:
    gheymat_taghir_sekke_hafte = 'black'
    

if week_color[4]==['low']:
    gheymat_taghir_gerami_hafte = 'red'
elif week_color[4]==['high']:
    gheymat_taghir_gerami_hafte = 'green'
else:
    gheymat_taghir_gerami_hafte = 'black'
    

if week_color[6]==['low']:
    gheymat_taghir_dollar_hafte = 'red'
elif week_color[6]==['high']:
    gheymat_taghir_dollar_hafte = 'green'
else:
    gheymat_taghir_dollar_hafte = 'black'





if day_color[0]==['low']:
    gheymat_taghir_ons = 'red'
elif day_color[0]==['high']:
    gheymat_taghir_ons = 'green'
else:
    gheymat_taghir_ons = 'black'
    
if day_color[1]==['low']:
    gheymat_taghir_sekke = 'red'
elif day_color[1]==['high']:
    gheymat_taghir_sekke = 'green'
else:
    gheymat_taghir_sekke = 'black'
    

if day_color[2]==['low']:
    gheymat_taghir_gerami = 'red'
elif day_color[2]==['high']:
    gheymat_taghir_gerami = 'green'
else:
    gheymat_taghir_gerami = 'black'
    

if day_color[3]==['low']:
    gheymat_taghir_dollar = 'red'
elif day_color[3]==['high']:
    gheymat_taghir_dollar = 'green'
else:
    gheymat_taghir_dollar = 'black'




workbook  = xlsxwriter.Workbook('price.xlsx')
worksheet = workbook.add_worksheet()

#caption = ('99/5/22')

red = workbook.add_format()
green = workbook.add_format()

red.set_font_color('red')
green.set_font_color('green')


worksheet.write('A2', 'انس طلا')
worksheet.write('A3', 'سکه امامی')
worksheet.write('A4', 'طلای 18 عیار')
worksheet.write('A5', 'دلارصرافی ملی')


worksheet.write('B1', 'قیمت امروز')
worksheet.write('C1', 'درصد تغییرات نسبت به روز قبل')
worksheet.write('D1', 'عملکرد هفته')
worksheet.write('E1', 'عملکرد هفته به درصد')





for row, row_data in enumerate(sep_data):
    worksheet.write_row(row + 1, 1, row_data)


if gheymat_taghir_ons=='red': #Change the price of auns up to date
    worksheet.write(1, 2,sep_data[0][1],red)
elif gheymat_taghir_ons == 'green':
    worksheet.write(1,2,sep_data[0][1],green)
else :
    worksheet.write(1,2,sep_data[0][1])
    

if gheymat_taghir_sekke =='red': #Change the price of coins up to date
    worksheet.write(2, 2,sep_data[1][1],red)
elif gheymat_taghir_sekke == 'green':
    worksheet.write(2,2,sep_data[1][1],green)
else :
    worksheet.write(2,2,sep_data[1][1])
    

if gheymat_taghir_gerami =='red': #Change the price of gold up to date
    worksheet.write(3, 2,sep_data[2][1],red)
elif gheymat_taghir_gerami == 'green':
    worksheet.write(3,2,sep_data[2][1],green)
else :
    worksheet.write(3,2,sep_data[2][1])
    

if gheymat_taghir_dollar=='red': #Change the price of doller up to date
    worksheet.write(4, 2,sep_data[3][1],red)
elif gheymat_taghir_dollar == 'green':
    worksheet.write(4,2,sep_data[3][1],green)
else :
    worksheet.write(4,2,sep_data[3][1])



if gheymat_taghir_ons_hafte=='red': #Changing the price of Auns per week
    worksheet.write(1, 3,sep_data[0][2],red)
    worksheet.write(1, 4,sep_data[0][3],red)

elif gheymat_taghir_ons_hafte == 'green':
    worksheet.write(1,3,sep_data[0][2],green)
    worksheet.write(1,4,sep_data[0][3],red)

else :
    worksheet.write(1,3,sep_data[0][2])
    worksheet.write(1, 4,sep_data[0][3])
 
    
if gheymat_taghir_sekke_hafte=='red': #Changing the price of coins per week
    worksheet.write(2, 3,sep_data[1][2],red)
    worksheet.write(2,4,sep_data[1][3],red)

elif gheymat_taghir_sekke_hafte== 'green':
    worksheet.write(2,3,sep_data[1][2],green)
    worksheet.write(2,4,sep_data[1][3],red)

else :
    worksheet.write(2,3,sep_data[1][2])
    worksheet.write(2, 4,sep_data[1][3])
  
    
if gheymat_taghir_gerami_hafte=='red': #Changing the price of gold per week
    worksheet.write(3, 3,sep_data[2][2],red)
    worksheet.write(3, 4,sep_data[2][3],red)

elif gheymat_taghir_gerami_hafte== 'green':
    worksheet.write(3,3,sep_data[2][2],green)
    worksheet.write(3, 4,sep_data[2][3],red)

else :
    worksheet.write(3,3,sep_data[2][2])
    worksheet.write(3, 4,sep_data[2][3])

    
if gheymat_taghir_dollar_hafte=='red': #Changing the price of doller per week
    worksheet.write(4, 3,sep_data[3][2],red)
    worksheet.write(4, 4,sep_data[3][3],red)

elif gheymat_taghir_dollar_hafte == 'green':
    worksheet.write(4,3,sep_data[3][2],green)
    worksheet.write(4, 4,sep_data[3][3],red)

else :
    worksheet.write(4,3,sep_data[3][2])
    worksheet.write(4, 4,sep_data[3][3])

    


workbook.close()

