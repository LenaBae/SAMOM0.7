#!/usr/bin/python3
# -*- coding:utf-8 -*-
#
# Copyright 2020 by JK Lee, JK Entertainment.
# All rights reserved.
# This file is part of the JK Entertainment Script Tool (JENST),
# and is released under the Python Software Foundation License. Please see the LICENSE
# file that should have been included as part of this package.
#
# 23-Jan-2021 by jaykay.lee 
# 25-Jan-2021 by Denny.Shin 
# since 19-Jan-2021 by Lena.bae 
# SA_MOM Initial version

import re
import os
from openpyxl import load_workbook 



#Color Scheme
GREY='\033[1;30m'
RED='\033[1;31m'
GREEN='\033[1;32m'
YELLOW='\033[0;33m'
BLUE='\033[1;34m'
NC='\033[0m'
Cyan='\033[1;36m'  
White='\033[1;37m'
pink='\033[1;35m'  

#files
#INPUT_TXT = 'Rowtxt.txt'
#INPUT_XSL = 'Rowdata.xlsx'
OUTPUT_FILE = 'output.xml'

def logo():
    print('{} _______  _______    _______  _______  _______  '.format(pink))
    print('|  _____||  ___  |  |       ||  ___  ||       |')
    print('| |      | |   | |  | || || || |   | || || || |')
    print('| |_____ | |___| |  | || || || |   | || || || |')
    print('|_____  ||  ___  |  | ||_|| || |   | || ||_|| |')
    print('      | || |   | |  | |   | || |   | || |   | |')
    print(' _____| || |   | |  | |   | || |___| || |   | |')
    print('|_______||_|   |_|  |_|   |_||_______||_|   |_| ')
    print('                                       {}ver. 0.5{}'.format(BLUE,NC))
    print('')
    print('                Make scf tool for SA deployment')
    print('')

def get_col(load_ws,col_name):
    value_list = list()
    for i, col in enumerate(load_ws.columns):
        if col[0].value == col_name:
            for j, cell in enumerate(col):
                if j == 0:
                    continue
                value_list.append(cell.value)
    return value_list

def main():
    os.system('COLOR 11') 
    logo()
    while True:
        try:
            INPUT_TXT=input('{}Please enter Parameter file ... {}'.format(White,NC))
            data = open(INPUT_TXT,'r')
            contents = data.read()
            data.close()
            break
        except:
            print('{}Can\'t find ({}){}'.format(RED,INPUT_TXT,NC))

    while True:
        try:
            INPUT_XSL=input('{}Please enter Local information file ... {}'.format(White,NC))
            load_wb = load_workbook(INPUT_XSL, data_only=True)
            load_ws = load_wb['Sheet1'] 
            break
        except:
            print('{}Can\'t find ({}){}'.format(RED,INPUT_XSL,NC))

    NRBTS_L = get_col(load_ws,'NRBTS')
    TAC_L = get_col(load_ws,'5G TAC')
    CELL_L = get_col(load_ws, 'NRCELL')
    header = re.findall('<?xml version=.*</header>\n',contents,re.DOTALL)
    maindata = re.findall('    <managedObject class=.*    </managedObject>\n',contents,re.DOTALL)
    last = re.findall('  </cmData>.*</raml>',contents,re.DOTALL)

    output_file = open(OUTPUT_FILE,'w') 
    output_file.write('<?')   
    output_file.write(str(*header))  
    for i, bts in enumerate(NRBTS_L):
        print('[{}] processing BTS({}), CELL({}), TAC({}) ... '.format(i, bts,CELL_L[i],TAC_L[i]), end='')
        changeBTS=re.sub('BTS-[0-9]+','BTS-'+str(bts),str(*maindata))
        changeCELL=re.sub('/NRCELL-[0-9]+','/NRCELL-'+str(CELL_L[i]),changeBTS)
        changeTac=re.sub('GsTac">[0-9]+','GsTac">'+str(TAC_L[i]),changeCELL)
        changeTrackingare=re.sub('AREA-[0-9]','AREA-'+str(TAC_L[i]),changeTac)
        output_file.write(changeTrackingare)
        print(' done')
    
    output_file.write(str(*last))

    output_file.close()
    input('Press Enter Key ...')



if __name__ == '__main__': 
    main()
