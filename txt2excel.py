#!/usr/bin/python
# -*- coding: UTF-8 -*-

import sys
import os
import re
import openpyxl

excel_dict = {}
# xls = {}
# sheet = {}

txt_temp = ''

def txt2excel(txt_file, name_prefix):
    abc = name_prefix.split('_')  #获取Excel的abc列内容
    print(abc)
    if len(abc) < 3:
        print('[Error] name of txt error')
        return
    
    pos = []  #保存txt表头每列的起始位置
    hijks = ["TestDescription/Event", "Result", "Value", "U.limit", "L.limit", "DUT", "Pin/pattern", "Sequence"]
    head = ["UUID(LOT_ID+WaferID)", "PROBE_XY", "Temperature(C)", "Power Voltage(V)", "Pin/Pattern", "LimitUpper", "LimitLower", "Result", "Test Condition1"]

    txtlines = txt_file.readlines()  #读取txt行数据
    for s in hijks:
        pos.append(txtlines[1].find(s))  #txt文件的第二行为表头，
    # print(pos)

    pos_len = len(pos)
    if (pos_len != len(hijks)):
        print('[Error] Head num error')
        return
    
    pos_range = []  #每一列的范围，按照Excel表的位置保存
    pos_range.append([pos[6], pos[7]])
    pos_range.append([pos[3], pos[4]])
    pos_range.append([pos[4], pos[5]])
    pos_range.append([pos[2], pos[3]])
    pos_range.append([pos[0], pos[1]])
    print(pos_range)

    xls = openpyxl.Workbook()

    excel_temp = ''  #判断Sequence是否与上一条不同
    global txt_temp  #判断是否是新的txt文件
    global excel_dict  #字典，保存每个Excel表的行位置

    start = False

    voltage = 0.0
    for k in range(len(txtlines)): #从头遍历txt文件

        txtline = txtlines[k]
        if txtline[0 : 6] == 'TestID':  #判断一个txt文件有多个测试
            # k += 2
            start = True
            continue
        elif txtline[0 : 6] == '** End':
            xls.save(excel_temp + '.xlsx')
            print('[Write] save excel:', excel_temp, k)
            voltage = 0.0
            start = False
        if not start:
            continue
        #判断此行是否为电压数值
        power = txtline[pos_range[-1][0] : pos_range[-1][1]]  
        if 'VccPowerFactor' in power:
            vcc = txtline[pos_range[3][0] : pos_range[3][1]]
            vcc = vcc.strip()
            voltage = float(vcc) * 3.0
            continue
        
        excelname = txtline[pos[-1]:-1]  #获取Sequence作为excel命名
        
        # print(excelname, k)
        if not excelname in excel_dict:  #如果无此表，保存上一次的结果，并另新建一个
            if os.path.exists(excelname + '.xlsx'):
                print('[Warn] please keep pwd no excel')
                return
            if excel_temp != '':  #判断是否为第一次有效行
                xls.save(excel_temp + '.xlsx')
                print('[Write] save excel:', excel_temp, k)
                xls = openpyxl.Workbook()

            print('[Write] new excel:', excelname, k)
            r = 2                #在excel开始写的位置（y）
            c = 1                #在excel开始写的位置（x）
            excel_dict[excelname] = r
            sheet = xls.create_sheet(excelname, 0)
            for i in range(len(head)):  #写表头数据
                sheet.cell(row = 1, column = i + 1, value = head[i])
        else:  #表已存在
            if txt_temp != '' and txt_temp != name_prefix:  #如果是后续文件
                if excel_temp != '' and excel_temp != excelname: #非首行出现上一条不同类型，保存之前结果
                    xls.save(excel_temp + '.xlsx')
                    print('[Write] save excel:', excel_temp, k)
            else:  #如果是第一个txt文件
                if excel_temp != excelname:
                    xls.save(excel_temp + '.xlsx')
                    print('[Write] save excel:', excel_temp, k)

            if excel_temp != excelname: #出现不同类型，则指向新类型
                    print('[Read] load excel:', excelname, k)
                    xls = openpyxl.load_workbook(excelname + '.xlsx')
                    sheet = xls.get_sheet_by_name(excelname)
        #上面判断主要目的指向正确的xls
        
        
        # sheet[excelname] = xls[excelname].get_sheet_by_name(excelname)
        #对每一列写数据
        for i in range(len(head)):
            if i < 3:
                sheet.cell(row = excel_dict[excelname], column = i + 1, value = abc[i])
            elif i == 3:
                sheet.cell(row = excel_dict[excelname], column = i + 1, value = voltage)
            elif i == 4: #不用考虑单位
                sheet.cell(row = excel_dict[excelname], column = i + 1, value = txtline[pos_range[i-4][0] : pos_range[i-4][1]])
            elif 4 < i < 8: #考虑单位
                txt_cell = txtline[pos_range[i-4][0] : pos_range[i-4][1]]
                pattern = r'[A-Za-z]+|[-]?\d+[.]?\d*'
                result = re.findall(pattern, txt_cell)
                if len(result) > 1:
                    sheet.cell(row = 1, column = i + 1, value = head[i] + '/' + result[1])
                    sheet.cell(row = excel_dict[excelname], column = i + 1, value = result[0])
                elif len(result) > 0:
                    sheet.cell(row = excel_dict[excelname], column = i + 1, value = result[0])
                else:
                    sheet.cell(row = excel_dict[excelname], column = i + 1, value = txt_cell)
            elif i == 8: #多条件处理
                event = txtline[pos_range[-1][0] : pos_range[-1][1]]
                if len(event.split(':')) <= 1:
                    continue
                desc = event.split(':')[0]
                date = event.split(':')[1] # date = ' 0x00; 0x01; -0uA'
                
                descs = desc.split('_')
                dates = date.split('_')
                if (len(descs) != len(dates)):
                    print('[Error] the num of events and values not same')
                    return
                # pattern = r'(.*?);'
                # result = re.findall(pattern, date) # result = {' 0x00', ' 0x01', ' -0uA'}
                if len(dates) > 1:  #多个条件
                    for j in range(len(dates)):
                        pattern1 = r'0[xX][0-9a-fA-F]+'  # 0x 十六进制
                        has_hex = re.findall(pattern1, dates[j])
                        if has_hex:
                            decimal = int(has_hex[0], 16)
                            test = descs[j]
                        else:  #十进制数值
                            pattern11 = r'[A-Za-z]+|[-]?\d+[.]?\d*'
                            has_unit = re.findall(pattern11, dates[j])
                            if len(has_unit) > 1: # 有单位
                                test = descs[j] + '/' + has_unit[1]
                                decimal = has_unit[0]
                            elif len(has_unit) > 0: # 无单位
                                test = descs[j]
                                decimal = has_unit[0]
                            else:
                                decimal = dates[j]
                        
                        sheet.cell(row = excel_dict[excelname], column = j + i + 1, value = decimal)
                        sheet.cell(row = 1, column = j + i + 1, value = test)
                else: # 只有一个条件
                    pattern2 = r'0[xX][0-9a-fA-F]+' #0x十六进制
                    result2 = re.findall(pattern2, date)
                    if result2:
                        test = descs[0]
                        sheet.cell(row = 1, column = i + 1, value = test)
                        decimal = int(result2[0], 16)
                    else:  #十进制数值
                        pattern21 = r'[A-Za-z]+|[-]?\d+[.]?\d*'
                        result21 = re.findall(pattern21, date)
                        if len(result21) > 1: # 有单位
                            test = descs[0] + '/' + result21[1]
                            sheet.cell(row = 1, column = i + 1, value = test)
                            decimal = result21[0]
                        elif len(result21) > 0: # 无单位
                            decimal = result21[0]
                        else:
                            decimal = date.strip()
                    
                    sheet.cell(row = excel_dict[excelname], column = i + 1, value = decimal)
                    # test = 'Test Condition' + str(j + 1)
                    # sheet.cell(row = 1, column = i + 1, value = test)

        excel_dict[excelname] += 1
        excel_temp = excelname
    txt_temp = name_prefix

def merger_excel():
    # IO1V8_XX IO3V3_XX前缀相同的合并到一个excel的不同sheet中
    prefix = {}  #字典，文件名前缀出现次数
    key = []  #列表，保存需要合并的文件的关键字
    filelist = []  #保存需要合并的文件
    listd = os.listdir('./')  #文件列表
    
    for file in listd:
        if file.endswith('.xlsx'):
            name = file.split('.xlsx')[0]
            name = name.split('_')[0]
            if not name in prefix:
                prefix[name] = 1
            else:
                prefix[name] += 1
    for k in prefix.keys():
        if prefix[k] > 1:
            key.append(k + '_')
    print(key)
    
    for i in range(len(key)):
        filelist_temp = []
        for file in listd:
            if key[i] in file:
                filelist_temp.append(file)
        filelist.append(filelist_temp)
    print(filelist)  #需要合并的文件列表

    for i in range(len(filelist)):
        if len(filelist[i]) == 1:
            print('[Warn] only one excel]')
            continue
        src = openpyxl.load_workbook(filelist[i][0])
        print('[Read] open ' + filelist[i][0])
        #取第一个文件，将剩下的文件复制到第一个文件的sheet
        for j in range(1, len(filelist[i])):
            target = openpyxl.load_workbook(filelist[i][j]).active
            target._parent = src
            src._add_sheet(target)
            sheet_name = filelist[i][j].split('.')[0]
            # target_sheet = target.get_sheet_by_name(sheet_name)
            # src.copy_worksheet(target_sheet)
            print('[Merger] sheet ' + sheet_name)
            os.remove(filelist[i][j])
            print('[Delete] sheet ' + sheet_name)
        ws = src['Sheet']
        src.remove(ws)
        src.save(filelist[i][0])
        test_name = ''
        #重命名处理
        if key[i] == 'IO1V8_' or key[i] == 'IO3V3_':
            test_name = 'DC_TEST'
        else:
            test_name = 'TEST'
        os.rename(filelist[i][0], key[i] + test_name + '.xlsx')
        print('[Name] new excel name ' + key[i] + 'DC_TEST' + '.xlsx')


if __name__ == "__main__":
    
    txt_num = len(sys.argv)
    txt_prefix = []  #保存txt文件前缀名
    merger = True
    for i in range(1, txt_num):
        if os.path.exists(sys.argv[i]):  #如果文件名存在
            txt_prefix.append(sys.argv[i].split('.txt')[0])
            print('[Read] txtFile %s : %s' % (i, sys.argv[i]))
        else:
            print('[Error] txtFile %s : %s not found' % (i, sys.argv[i]))
            sys.exit()
    for i in range(1, txt_num):
        txtfile = open(sys.argv[i], 'r')
        txt2excel(txtfile, txt_prefix[i - 1])
        txtfile.close()
    if merger:  #合并文件
        merger_excel()
        
