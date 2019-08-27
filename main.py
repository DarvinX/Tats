import PySimpleGUI as sg
from datetime import datetime
from comm import comm as cm

size = (20,2)
Type = ['Print', 'Photo Print', 'Scan']
PrintType = ['Monochrome', 'Colored', 'Blank']
Columns = ['Name', 'Blank Sheet', 'Print Colored', 'Print B/W', 'Print Photo', 'Xerox Colored', 'Xerox B/W', 'Scan']


layout = [[sg.Text('Name: ', size=size), sg.InputText()],
            [sg.Text('Work Type'),sg.Combo(Columns[1:])],
            [sg.Text('Copies: ', size=size), sg.InputText(1)],
            [sg.Text('Note(optional): ', size=size), sg.InputText()],
            [sg.Text('Time: ', size=size), sg.Text(str(datetime.now()))],
            [sg.Submit(), sg.Exit()]
        ]

window = sg.Window('Tats', layout, default_element_size=(30,1), grab_anywhere=False,resizable=True, font=('Helvetica', 17))

event, values = window.Read()

if event not in [None, 'Exit']:
    com = cm('file.xls')