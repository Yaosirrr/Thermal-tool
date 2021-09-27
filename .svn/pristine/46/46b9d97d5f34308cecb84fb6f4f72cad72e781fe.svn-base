"""
Date: 2020/11/11
Author: Jeffrey
Last Modified: 2021/01/23
"""

from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
import re
from os import listdir
import argparse


TEMPLATE = 'Thermal.xlsx'
pattern1 = r'\d+'
pattern2 = r'\d+\.\d+'
get_data_methods = {}


def register(key):
    def inner(func):
        get_data_methods[key] = func
    return inner


def get_parser():
    parser = argparse.ArgumentParser(description="parameters list")
    parser.add_argument('--output', required=True)
    parser.add_argument('--hdd', required=False)
    parser.add_argument('--bmc', required=False)
    parser.add_argument('--bmc2', required=False)
    parser.add_argument('--ptu', required=False)
    parser.add_argument('--ptu2', required=False)
    parser.add_argument('--cpulog', required=False)
    parser.add_argument('--gpulog', required=False)
    parser.add_argument('--draw', required=False, choices=['enable', 'disable'], default='enable')
    return parser


def get_requirement(sheetname):
    wb = load_workbook(TEMPLATE)
    ws = wb[sheetname]
    requirement = None

    if sheetname == 'PTU':
        # { CPU0: [info], CPU1: [info], CPU2[info]}
        requirement = {}
        for cell in list(ws.rows)[1:]:
            if not cell[0].value:
                continue
            if cell[0].value.upper() not in requirement:
                requirement[cell[0].value.upper()] = []
            requirement[cell[0].value.upper()].append(cell[1].value.upper())

    elif sheetname == 'sensor':
        requirement = []
        # ignore capitals, [nlet_Temp, Outlet_Temp, CPU0_DTS]
        for cell in list(ws.rows)[1:]:
            if not cell[0].value:
                continue
            requirement.append(cell[0].value.upper())

    elif sheetname == 'HDD':
        requirement = []
        for cell in list(ws.rows)[1:]:
            if not cell[0].value:
                continue
            requirement.append(cell[0].value.lower())
    
    elif sheetname == 'BMC_GPU':
        # { CPU0: [info], CPU1: [info], CPU2[info]}
        requirement1 = [row[0].value for row in list(ws.rows)[1:] if row[0].value]
        nickname = [row[1].value for row in list(ws.rows)[1:] if row[1].value]
        requirement = [requirement1, nickname]
    
    elif sheetname == 'BMC_CPU':
        # { CPU0: [info], CPU1: [info], CPU2[info]}
        requirement = [[], []]
        for row in list(ws.rows)[1:]:
            if row[0].value:
                requirement[0].append(row[0].value)
            if row[1].value:
                requirement[1].append(row[1].value)

    return requirement


def data_analysis(row, col, ws, ws2):
    ws2.cell(row=2, column=1, value='max')
    ws2.cell(row=3, column=1, value='min')
    ws2.cell(row=4, column=1, value='3/4 average')
    ws2.cell(row=5, column=1, value='last 10 average')
    for c in range(2, col):
        data = []
        for r in range(2,row):
            if isinstance(a := ws.cell(r, c).value, (int, float)):
                data.append(a)
            else:
                print(f'Warning: data missing at log sheet: row {r}, col {c}. \
                    \n         Summary sheet returns nothing, please check')
                break
        else:
            # name = ws.cell(row=1, column=c).value
            maximum = max(data) if data else 0
            ws2.cell(row=2, column=c, value=maximum)
            minimum = min(data) if data else 0
            ws2.cell(row=3, column=c, value=minimum)
            avg = sum(data[:int(len(data)*3/4)])/int(len(data)*3/4) if data else 0
            ws2.cell(row=4, column=c, value=round(avg, 2))
            avg_last_10 = sum(data[-10:])/10 if data else 0
            ws2.cell(row=5, column=c, value=round(avg_last_10, 2))


def draw(ws1, ws2, row, col, title, x_title, y_title, start_row=1, start_col=2):
    #  draw
    c1 = LineChart()
    c1.title = title
    c1.style = 13
    c1.y_axis.title = y_title
    c1.x_axis.title = x_title
    data = Reference(ws1, min_col=start_col, min_row=start_row, max_col=col-1, max_row=row-1)
    c1.add_data(data, titles_from_data=True)
    for s in c1.series:
        s.smooth = True
    ws2.add_chart(c1, "A" + str(6))


@register('ptu2')
def get_ptu_data2(file_name, draw_pic):

    requirement = get_requirement('PTU')
    result = {}  # { Index: { CPU0: [info], CPU1: [info], CPU2: [info] } }
    start = False
    title = []
    with open(file_name) as f:

        for line in f.readlines():
            # ignore the first few lines, until find Time
            line = line.strip()
            if not start:
                if line.startswith('Index'):
                    start = True
                    title = re.split(r' +', line.upper())
            # start to access data { Index: { CPU0: [info], CPU1: [info], CPU2: [info] } }
            elif start:
                if line.startswith('Command:'):
                    break
                if re.match(r'\d+', line):
                    data = re.split(r' +', line)

                    if data[0] not in result.keys():
                        result[data[0]] = {}  # { time: {} }
                    device = data[title.index('DEVICE')]
                    if device not in requirement.keys():
                        continue
                    info = [data[title.index(requirement[device][i])] for i in range(len(requirement[device]))]
                    result[data[0]][device] = info  # {Index: {CPU0: [info]}}
    wb = load_workbook(TEMPLATE)
    
    if 'PTU log' in wb.sheetnames:
        wb.remove(wb['PTU log'])
    ws = wb.create_sheet('PTU log')


    # generate table title
    col = 2
    for key, val in requirement.items():
        for i in val:
            ws.cell(row=1, column=col, value=key + '_' + i)
            col += 1

    # fill in data
    row = 2
    for time, data in result.items():
        col = 2
        ws.cell(row=row, column=1, value=time)
        for info in data.values():
            for piece in info:
                if re.match(pattern2,piece):
                    ws.cell(row=row, column=col, value=float(piece))
                elif re.match(pattern1,piece):
                    ws.cell(row=row, column=col, value=int(piece))
                else:
                    ws.cell(row=row, column=col, value=piece)
                col += 1
        row += 1

    if 'PTU summary' in wb.sheetnames:
        wb.remove(wb['PTU summary'])
    ws2 = wb.create_sheet('PTU summary')

    # generate table title
    ccol = 2
    for key, val in requirement.items():
        for i in val:
            ws2.cell(row=1, column=ccol, value=key + '_' + i)
            ccol += 1

    # analyse data
    if draw_pic == 'enable':
        try:
            data_analysis(row, col, ws, ws2)
        except Exception:
            print(f'{ws.title} fails at max, min, average')
        draw(ws, ws2, row, col, 'Summary', 'Time', 'Value')
    wb.save(TEMPLATE)

    print('PTU2 DONE!')

    return result


@register('ptu')
def get_ptu_data(file_name, draw_pic):

    requirement = get_requirement('PTU')
    result = {}  # { time: { CPU0: [info], CPU1: [info], CPU2: [info] } }
    start = False
    title = []
    with open(file_name) as f:

        for line in f.readlines():
            # ignore the first few lines, until find Time
            if not start:
                if line.startswith('Time'):
                    start = True
                    title = re.split(r' +', line.upper())

            # start to access data { time: { CPU0: [info], CPU1: [info], CPU2: [info] } }
            elif start:
                if re.match(r'\d+\.\d+_\d+', line):
                    data = re.split(r' +', line)
                    # transfer '155212.909_1' to '15:52:12'
                    a = re.search(r'\d+', data[0]).group()
                    a = list(a)
                    a.insert(-2, ':')
                    a.insert(-5, ':')
                    data[0] = ''.join(a)
                    if data[0] not in result.keys():
                        result[data[0]] = {}  # { time: {} }
                    device = data[title.index('DEV')]
                    if device not in requirement.keys():
                        continue
                    info = [data[title.index(requirement[device][i])] for i in range(len(requirement[device]))]
                    result[data[0]][device] = info  # {time: {CPU0: [info]}}
                    
    wb = load_workbook(TEMPLATE)
    
    if 'PTU log' in wb.sheetnames:
        wb.remove(wb['PTU log'])
    ws = wb.create_sheet('PTU log')


    # generate table title
    col = 2
    for key, val in requirement.items():
        for i in val:
            ws.cell(row=1, column=col, value=key + '_' + i)
            col += 1

    # fill in data
    row = 2
    for time, data in result.items():
        col = 2
        ws.cell(row=row, column=1, value=time)
        for info in data.values():
            for piece in info:
                if re.match(pattern2,piece):
                    ws.cell(row=row, column=col, value=float(piece))
                elif re.match(pattern1,piece):
                    ws.cell(row=row, column=col, value=int(piece))
                else:
                    ws.cell(row=row, column=col, value=piece)
                col += 1
        row += 1

    if 'PTU summary' in wb.sheetnames:
        wb.remove(wb['PTU summary'])
    ws2 = wb.create_sheet('PTU summary')

    # generate table title
    ccol = 2
    for key, val in requirement.items():
        for i in val:
            ws2.cell(row=1, column=ccol, value=key + '_' + i)
            ccol += 1

    # analyse data
    if draw_pic == 'enable':
        try:
            data_analysis(row, col, ws, ws2)
        except Exception:
            print(f'{ws.title} fails at max, min, average')
        draw(ws, ws2, row, col, 'Summary', 'Time', 'Value')
    wb.save(TEMPLATE)

    print('PTU DONE!')

    return result


@register('bmc2')
def get_bmc_data2(file_name, draw_pic):
    requirement = get_requirement('sensor')
    result = {}  # { key: [value], ... }

    with open(file_name) as f:

        for line in f.readlines():
            title = re.search(r'\S+', line).group() if re.search(r'\S+', line) else ''
            if title.upper() in requirement:
                if title.upper() in result.keys():
                    value = re.search(r'.+\|\s(\d+).+\|.+', line).group(1) \
                        if re.search(r'.+\|\s(\d+).+\|.+', line) else 'disabled'
                    result[title.upper()].append(value)
                else:
                    result[title.upper()] = []

    wb = load_workbook(TEMPLATE)

    if 'sensor log' in wb.sheetnames:
        wb.remove(wb['sensor log'])
    ws = wb.create_sheet('sensor log')

    # generate table title
    col = 2
    for i in requirement:
        ws.cell(row=1, column=col, value=i)
        col += 1

    # fill in data
    col, row = 2, 2
    for key, val in result.items():
        row = 2
        for info in val:
            if re.match(pattern2, info):
                ws.cell(row=row, column=col, value=float(info))
            elif re.match(pattern1, info):
                ws.cell(row=row, column=col, value=int(info))
            else:
                ws.cell(row=row, column=col, value=info)
            row += 1
        col += 1

    if 'sensor summary' in wb.sheetnames:
        wb.remove(wb['sensor summary'])
    ws2 = wb.create_sheet('sensor summary')

    # generate table title
    ccol = 2
    for i in requirement:
        ws2.cell(row=1, column=ccol, value=i)
        ccol += 1

    # analyse data
    if draw_pic == 'enable':
        try:
            data_analysis(row, col, ws, ws2)
        except Exception:
            print(f'{ws.title} fails at max, min ,average')
        draw(ws, ws2, row, col, 'Summary', 'Time', 'Value')
    wb.save(TEMPLATE)

    print('BMC DONE!')

    return result


@register('bmc')
def get_bmc_data(file_name, draw_pic):

    requirement = get_requirement('sensor')
    result = {}  # { time: { data_key: data_value, ... }, }
    time = r'\d+:\d+:\d+'
    start = False

    with open(file_name) as f:
        for line in f.readlines():
            if not start:
                if re.match(r'\*+(Temp|Switch)\*+', line):
                    tem_result = {}
                    start = True
                    continue
            elif start:
                if a := re.search(time, line):
                    result[a.group()] = tem_result
                    start = False
                else:
                    data = re.split(r' +\| +', line)
                    # ['Inlet_Temp', '00h', 'ok', '55.0', '32 degrees C']
                    if data[0].upper() in requirement:
                        try:
                            tem_result[data[0].upper()] = \
                                data[4].rstrip().replace(' degrees C', '').replace(' Volts', '')\
                                .replace(' Watts','').replace(' RPM','')
                        except IndexError:
                            tem_result[data[0].upper()] = ''
    wb = load_workbook(TEMPLATE)
    
    if 'sensor log' in wb.sheetnames:
        wb.remove(wb['sensor log'])
    ws = wb.create_sheet('sensor log')

    # generate table title
    col = 2
    for i in requirement:
        ws.cell(row=1, column=col, value=i)
        col += 1

    # fill in data
    row = 2
    for time, data in result.items():
        col = 2
        ws.cell(row=row, column=1, value=time)
        for info in data.values():
            if re.match(pattern2,info):
                ws.cell(row=row, column=col, value=float(info))
            elif re.match(pattern1,info):
                ws.cell(row=row, column=col, value=int(info))
            else:
                ws.cell(row=row, column=col, value=info)
            col += 1
        row += 1

    if 'sensor summary' in wb.sheetnames:
        wb.remove(wb['sensor summary'])
    ws2 = wb.create_sheet('sensor summary')

    # generate table title
    ccol = 2
    for i in requirement:
        ws2.cell(row=1, column=ccol, value=i)
        ccol += 1

    # analyse data
    if draw_pic == 'enable':
        try:
            data_analysis(row, col, ws, ws2)
        except Exception:
            print(f'{ws.title} fails at max, min ,average')
        draw(ws, ws2, row, col, 'Summary', 'Time', 'Value')
    wb.save(TEMPLATE)

    print('BMC DONE!')

    return result


@register('hdd')
def get_hdd_data(file_name, draw_pic):
    requirement = get_requirement('HDD')
    result = {}  # { time: { data_key: data_value, ... }, }
    time = r'\d+:\d+:\d+'

    with open(file_name) as f:

        for line in f.readlines():
            if line.startswith('DEV'):
                tem_result = {}
                continue
            elif a := re.search(time, line):
                result[a.group()] = tem_result
            else:
                data = re.split(r' +', line)
                # ['Inlet_Temp', '00h', 'ok', '55.0', '32 degrees C']
                if data[0].lower() in requirement:
                    try:
                        tem_result[data[0].lower()] = data[1]
                    except IndexError:
                        tem_result[data[0].lower()] = ''

    wb = load_workbook(TEMPLATE)
    
    if 'hdd log' in wb.sheetnames:
        wb.remove(wb['hdd log'])
    ws = wb.create_sheet('hdd log')

    # generate table title
    col = 2
    for i in requirement:
        ws.cell(row=1, column=col, value=i)
        col += 1

    # fill in data
    row = 2
    for time, data in result.items():
        col = 2
        ws.cell(row=row, column=1, value=time)
        for info in data.values():
            if re.match(pattern2,info):
                ws.cell(row=row, column=col, value=float(info))
            elif re.match(pattern1,info):
                ws.cell(row=row, column=col, value=int(info))
            else:
                ws.cell(row=row, column=col, value=info)
            col += 1
        row += 1

    if 'hdd summary' in wb.sheetnames:
        wb.remove(wb['hdd summary'])
    ws2 = wb.create_sheet('hdd summary')

    # generate table title
    ccol = 2
    for i in requirement:
        ws2.cell(row=1, column=ccol, value=i)
        ccol += 1

    # analyse data
    if draw_pic == 'enable':
        try:
            data_analysis(row, col, ws, ws2)
        except Exception:
            print(f'{ws.title} fails at max, min ,average')
        draw(ws, ws2, row, col, 'Summary', 'Time', 'Value')
    wb.save(TEMPLATE)

    print('HDD DONE!')
    return result


@register('gpulog')
def get_gpu_data(file_name, draw_pic):
    """
    :param file_name:
    :return: [ { data1 }, { data2 }, { data3 }, { data4 } ], for each data, obeys the key: [ value1, value2 ] structure
    """
    [requirement, nickname] = get_requirement('BMC_GPU')

    result = []
    Final_Domain_Output_Duty = {'PWM0':[], 'PWM1':[], 'PWM2':[], 'PWM3':[], 'PWM4':[]}
    start = False
    change_line = False
    with open(file_name) as f:
        previous_result = {}  # initialize the container
        try:
            for line in f.readlines():
                if line.startswith('Timer'):
                    start = True  # ignore the first few lines
                    if previous_result:
                        result.append(previous_result)  # store the previous data
                    previous_result = {}  # empty the container
                # start to access data
                elif start:
                    if not line:
                        # blank
                        continue
                    elif line.startswith('[Final Domain Output Duty]') or change_line:
                        if change_line:
                            data = re.split(r'\s\|\s', line.rstrip(' |\n'))
                            # ['PWM0 = 55', 'PWM1 = 55', 'PWM2 = 74', 'PWM3 = 74', 'PWM4 = 74']
                            data = dict(zip(Final_Domain_Output_Duty, data))
                            for key, val in Final_Domain_Output_Duty.items():
                                if key in data:
                                    val.append(int(re.search(r'\w+\s=\s(\d+)', data[key]).group(1)))
                                else:
                                    val.append(None)
                        change_line = not change_line

                    elif line.startswith('+'):
                        # case1
                        data = re.split(r'\s*=\s*', line.lstrip('+'))
                        key = data[0]
                        if key in requirement:
                            previous_result[key] = [re.sub(r'\D+', '', data[1])]
                    elif line.startswith('['):
                        # case2
                        data = re.split(r', ', line)
                        key = data[0]
                        if key in requirement and len(data) > 3:
                            temp = re.sub(r'\D+', '', data[1])
                            pwm = re.sub(r'\D+', '', data[3])
                            previous_result[key] = [temp, pwm]
                    elif '|' in line:
                        # PWM0 = 32 | PWM1 = 32 | PWM2 = 32 | PWM3 = 32 | PWM4 = 32 |
                        continue
                    else:
                        # case3
                        data = re.split(r' *= *', line)
                        key = re.sub(r' *\(.+\)', '', data[0])
                        if key in requirement and len(data) > 1:
                            previous_result[key] = [re.sub(r'\D+', '', data[1])]
        except Exception as e:
            print(f'Error at [{line}], please check your data. {e}')
        else:
            print('get GPU data success, writing to excel...')
    wb = load_workbook(TEMPLATE)

    if 'BMC_GPU log' in wb.sheetnames:
        wb.remove(wb['BMC_GPU log'])
    ws = wb.create_sheet('BMC_GPU log')

    # generate table title
    col = 2
    for i in range(len(requirement)):
        if str(requirement[i]).startswith('['):
            ws.cell(row=1, column=col, value=str(nickname[i])+'_Temp')
            col += 1
            ws.cell(row=1, column=col, value=str(nickname[i])+'_PWM')
            col += 1
        else:
            ws.cell(row=1, column=col, value=str(nickname[i]))
            col += 1
    for i in range(5):
        ws.cell(row=1, column=col, value='Final PWM'+str(i))
        col += 1

    # fill in data
    col, row = 2, 2
    for each_dict_data in result:
        col = 2
        for key in requirement:
            try:
                for value in each_dict_data[key]:
                    value = transfer_str_2_num(value)
                    ws.cell(row=row, column=col, value=value)
                    col += 1
            except KeyError:
                print(f'Warning: {key} missing at row {row} column {col}.')
                ws.cell(row=row, column=col, value=None)
                if str(key).startswith('['):
                    col += 2
                else:
                    col += 1
        ws.cell(row=row, column=1, value=row-2)
        row += 1

    new_row, new_col = 2, col
    for key, vals in Final_Domain_Output_Duty.items():
        new_row = 2
        for val in vals:
            ws.cell(row = new_row, column = new_col, value = val)
            new_row += 1
        new_col += 1


    if 'BMC_GPU summary' in wb.sheetnames:
        wb.remove(wb['BMC_GPU summary'])
    ws2 = wb.create_sheet('BMC_GPU summary')

    if draw_pic == 'enable':
        draw(ws, ws2, row, col, 'GPU Summary', 'Time', 'Value',1,2)
    wb.save(TEMPLATE)

    print('write GPU success!')
    return result


@register('cpulog')
def get_cpu_data(file_name, draw_pic):
    '''
    Id:0,sensorindex:15;CPU0_DTS_MARGIN_TEMP
    Final pwm:pwmnum--0
    Final 8056 pwm:pwmnum--2
    :return:
    '''
    [requirement, nickname] = get_requirement('BMC_CPU')

    result = []
    start = False
    with open(file_name) as f:
        previous_result = {}  # initialize the container
        try:
            for line in f.readlines():
                if line.startswith('xxx-test==='):
                    start = True  # ignore the first few lines
                    if previous_result:
                        result.append(previous_result)  # store the previous data
                    previous_result = {}  # empty the container
                # start to access data
                elif start:
                    if not line:
                        # blank
                        continue
                    elif line.startswith('Id'):
                        data = re.split(r' +', line)
                        key = data[0]
                        if key in requirement:
                            # data
                            temp = re.search(r'temp(:| )\d+\.\d+', line).group() if re.search(r'temp(:| )\d+\.\d+', line) else ''
                            pwm = re.search(r'pwm:\d+\.\d+', line).group() if re.search(r'pwm:\d+\.\d+', line) else ''
                            previous_result[key] = [re.sub(r'temp(:| )', '', temp), re.sub(r'pwm:', '', pwm)]
                    elif line.startswith(' pid:'):
                        temp = re.search(r'temp:\d+\.\d+', line).group()
                        pwm = re.search(r'pwm:\d+\.\d+', line).group()
                        previous_result[key] = [re.sub(r'temp(:| )', '', temp), re.sub(r'pwm:', '', pwm)]
                    else:
                        data = re.split(r': +', line)
                        key = data[0]
                        if key in requirement:
                            # data
                            rotate = re.search(r': \d+ ', line).group() if re.search(r': \d+ ', line) else ''
                            previous_result[key] = [re.sub(r'\D+', '', rotate)]
        except Exception:
            print(f'Error at [{line}], please check your data.')
        else:
            print('get CPU data success, writing to excel...')
    wb = load_workbook(TEMPLATE)

    if 'BMC_CPU log' in wb.sheetnames:
        wb.remove(wb['BMC_CPU log'])
    ws = wb.create_sheet('BMC_CPU log')

    # generate table title
    col = 4
    for i in range(len(requirement)):
        if str(requirement[i]).startswith('Id'):
            ws.cell(row=2, column=col, value=str(nickname[i])+'_Temp')
            col += 1
            ws.cell(row=2, column=col, value=str(nickname[i])+'_PWM')
            col += 1
        else:
            ws.cell(row=2, column=col, value=str(nickname[i]))
            col += 1

    # fill in data
    col, row = 4, 3
    for each_dict_data in result:
        col = 4
        for key in requirement:
            try:
                for value in each_dict_data[key]:
                    value = transfer_str_2_num(value)
                    ws.cell(row=row, column=col, value=value)
                    col += 1
            except KeyError:
                print(f'Warning: {key} missing at row {row} column {col}.')
                ws.cell(row=row, column=col, value=None)
                if str(key).startswith('Id'):
                    col += 2
                else:
                    col += 1
        ws.cell(row=row, column=3, value=row - 3)
        row += 1

    if 'BMC_CPU summary' in wb.sheetnames:
        wb.remove(wb['BMC_CPU summary'])
    ws2 = wb.create_sheet('BMC_CPU summary')

    if draw_pic == 'enable':
        draw(ws, ws2, row, col, 'CPU Summary', 'Time', 'Value',2,4)
        
    # change title
    for cell in list(ws.rows)[1]:
        if str(cell.value).startswith('Id'):
            cell.value = cell.value.split(';')[-1]

    wb.save(TEMPLATE)

    print('write CPU success!')


def transfer_str_2_num(input_str):
    _format = [r'\d+\.\d\d', r'\d+\.\d\d\d\d\d\d', r'\d+']
    if input_str:
        if re.match(_format[0], input_str):
            return float(input_str)
        elif re.match(_format[1], input_str):
            return float(input_str[:-4])
        else:
            return int(float(input_str))
    else:
        return None


def main():
    
    parser = get_parser()
    args = parser.parse_args()
    cli_args = dict(vars(args))
    global TEMPLATE
    TEMPLATE = args.output
    for method, file in cli_args.items():
        if not file or method not in get_data_methods:
            continue
        try:
            print(f'start to process {file}...')
            get_data_methods[method](file, cli_args['draw'])
        except Exception as e:
            print(f'Fail in {file}.\nDetails: {e}')


if __name__ == '__main__':
    main()
