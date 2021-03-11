# 23-Jan-2021 by jaykay.lee 
# 25-Jan-2021 by Denny.Shin 
# since 19-Jan-2021 by Lena.bae 
# SA_MOM Initial version

import re
import os
from openpyxl import load_workbook 

#Color_Scheme
GREY='\033[1;30m'
RED='\033[1;31m'
GREEN='\033[1;32m'
YELLOW='\033[0;33m'
BLUE='\033[1;34m'
NC='\033[0m'
Cyan='\033[1;36m'  
White='\033[1;37m'
pink='\033[1;35m'  


#SA_MOM_logo
def logo():
    print('{} _______  _______    _______  _______  _______  '.format(pink))
    print('|  _____||  ___  |  |       ||  ___  ||       |')
    print('| |      | |   | |  | || || || |   | || || || |')
    print('| |_____ | |___| |  | || || || |   | || || || |')
    print('|_____  ||  ___  |  | ||_|| || |   | || ||_|| |')
    print('      | || |   | |  | |   | || |   | || |   | |')
    print(' _____| || |   | |  | |   | || |___| || |   | |')
    print('|_______||_|   |_|  |_|   |_||_______||_|   |_| ')
    print('                                       {}ver. 0.7{}'.format(BLUE,NC))
    print('                                    {}by Lena.Bae{}'.format(BLUE,NC))
    print('')

#load_ws
def get_col(load_ws,col_name):
    value_list = list()
    for i, col in enumerate(load_ws.columns):
        if col[0].value == col_name:
            for j, cell in enumerate(col):
                if j == 0:
                    continue
                value_list.append(cell.value)
    return value_list

#file_open_read
def main():
    os.system('COLOR 11')
    logo()
    while True:
        #텍스트 파일 요청
        try:
            INPUT_TXT=input('{}Please enter Parameter file ... {}'.format(White,NC))
            data = open(INPUT_TXT,'r')
            contents = data.read()
            data.close()
            break
        #텍스트 파일 없을시 다시 요청
        except:
            print('{}Can\'t find ({}){}'.format(RED,INPUT_TXT,NC))

    while True:
        #엑셀 파일 요청
        try:
            INPUT_XSL=input('{}Please enter Local information file ... {}'.format(White,NC))
            load_wb = load_workbook(INPUT_XSL, data_only=True)
            load_ws = load_wb['Sheet1'] 
            sheet1 = load_wb.active
            break
        #엑셀 파일 없을시 다시 요청
        except:
            print('{}Can\'t find ({}){}'.format(RED,INPUT_XSL,NC))

    i = 0
    xlsxnameList=[]#output 생성을 위한 열이름 List
    xlsxvalue=[]#output 생성을 위한 열이름에 해당하는값 List
   
    while True:
        xlsxname = sheet1[chr(65+i)+'1'].value #입력한 엑셀의 열이름 List
   
        if xlsxname == None: #엑셀의 열이름 없을때 아래 실행          
            break

        else:#xlsxnameList에 열이름 추가
            if xlsxname == "TRACKINGAREA-":#TRACKINGAREA- 는 create가 있을때만 실행하기위해, operation="create"를 찾음 .*에 create 넣어도 가능 
                TRACKCREATE = re.findall('" operation=".*"',str(re.findall('TRACKINGAREA-[0-9].*operation=".*"',contents)[0]))
                xlsxnameList.append(re.findall(xlsxname,contents)[0])
            elif xlsxname == "NRBTS-": #NRBTS-로 입력받은 값을 BTS-로 인식하여 MRBTS- 도 바꾸게함
                xlsxnameList.append(re.findall("BTS-",contents)[0])
            else:#특이사항 없는 엑셀 열이름 그대로 xlsxnameList에 넣음
                xlsxnameList.append(re.findall(xlsxname,contents)[0])

            xlsxvalue.append(get_col(load_ws,xlsxname))
            i=i+1   
    #<?xml version 부터 </header>까지 header 부분으로 지정          
    header = re.findall('<?xml version=.*</header>',contents,re.DOTALL)

    #<managedObject class= 부터 </raml>까지 maindata로 지정, output에서 반복되는 구간
    maindata = re.sub('''
  </cmData>
</raml>''','',str(*re.findall('\n    <managedObject class=.*</raml>',contents,re.DOTALL)))

    #마지막 부분지정   
    last = re.findall('''
  </cmData>
</raml>''',contents,re.DOTALL)

    #output.xml 생성
    output_file = open('output.xml','w') 
    output_file.write('<?')   
    output_file.write(str(*header))#header 입력
    for i, value in enumerate(xlsxvalue[0]):
        print('\n',i,'processing...\n/',end='')
        for a, name in enumerate(xlsxnameList): 
            print(xlsxnameList[a],xlsxvalue[a][i],'/',end='')

            #maindata 변경후 입력
            if name == 'TRACKINGAREA-':
                #create 포함한 TRACKINGAREA값을 xlsxvalue값으로 변경,TRACKCREATE[0]="operation="create"
                maindata=re.sub(name+'.*'+str(TRACKCREATE[0]),name+str(xlsxvalue[a][i])+str(TRACKCREATE[0]),maindata)

                #</p> 포함한 TRACKINGAREA변경
                maindata=re.sub(name+'[\d]</p>',name+str(xlsxvalue[a][i])+'</p>',maindata)

            else:#숫자,true,false를 xlsxvalue값으로 변경
                maindata=re.sub(name+'[\d|true|false]+',name+str(xlsxvalue[a][i]),maindata)

        output_file.write(maindata)
    print('\ndone')
    #last 입력
    output_file.write(str(*last))

    output_file.close()
    input('Press Enter Key ...')

if __name__ == '__main__': 
    main()