import PySimpleGUI as sg
from datetime import datetime
import xlwt
import xlrd
from xlutils.copy import copy

size = (20,2)
Columns = ['Name', 'Time', 'Type', 'Print Type', 'Copies', 'Note']
Type = ['Print', 'Photo Print', 'Scan']
PrintType = ['Monochrome', 'Colored', 'Blank']
style1 = xlwt.easyxf(num_format_str='DD/MM/YYYY hh:mm:ss')
headingStyle = xlwt.easyxf('font: name Times New Roman, bold on')

layout = [[sg.Text('Name: ', size=size), sg.InputText()],
            [sg.Text('Type: ', size=size), sg.Radio('Print', 'Radio02', default=True), sg.Radio('Photo Print', 'Radio02'), sg.Radio('Scan', 'Radio02')],
            [sg.Text('Print Type: ', size=size), sg.Radio('Monochrome', 'Radio01'), sg.Radio('Colored', 'Radio01'), sg.Radio('Blank', 'Radio01')],
            [sg.Text('Copies: ', size=size), sg.InputText(1)],
            [sg.Text('Note(optional): ', size=size), sg.InputText()],
            [sg.Text('Time: ', size=size), sg.Text(str(datetime.now()))],
            [sg.Submit(), sg.Exit()]
        ]

window = sg.Window('Tats', layout, default_element_size=(30,1), grab_anywhere=False,resizable=True, font=('Helvetica', 17))

event, values = window.Read()

if event not in [None, 'Exit']:
    try:
        rb = xlrd.open_workbook('info.xls', formatting_info=True)
        r_sheet = rb.sheet_by_index(0)
        r = r_sheet.nrows
        wb = copy(rb)
        sheet = wb.get_sheet(0)
    except FileNotFoundError:
        print('file not found\n')
        print('creating file\n')
        r=1
        wb = xlwt.Workbook()
        sheet = wb.add_sheet('sheet1')
        for i in range(len(Columns)):
            sheet.write(0, i, Columns[i], headingStyle)
    except:
        print('something went wrong')

    sheet.write(r, 0, values[0])
    sheet.write(r, 1, datetime.now(), style1)
    for i in [1,2,3]:
        if values[i]:
            sheet.write(r, 2, Type[i-1])
            break
    if not values[3]:
        for i in [4,5,6]:
            if values[i]:
                sheet.write(r, 3, PrintType[i-4])
                break

    sheet.write(r, 4, values[7])
    sheet.write(r, 5, values[8])
    wb.save('info.xls')