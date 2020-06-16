import sys
import PySimpleGUI as sg 
import subprocess
import os
import d3dshot

def Launcher():
    def print(line):
        window['output'].update(line)
    sg.theme('black')

    sg.set_options(element_padding=(1,1), button_element_size=(10,1), auto_size_buttons=False)

    col1 = [[sg.Text('Save Options', font='Calibri 8')],
            [sg.Text(' '), sg.Text('Filename: ', text_color='blue'), sg.Input(size=(12, 1), key='-vFilename-')],
            [sg.Text(' ' * 5), sg.Text('Folder: ', text_color='blue'), sg.Input(size=(12,1), key='-vFolder-'), 
            sg.FolderBrowse(button_color=('white', 'black'), font='Calibri 10', size=(7,1))],
            [sg.Text(' ', size=(30,1), font='Calibri 8', text_color='red', key='output')],]

    layout = [[sg.Text('QuickSnip', text_color='blue', font='Forte 14'), sg.Text(' ' * 60), sg.Button('x', size=(4,1), button_color=('white', '#fc0303'))],
                [sg.Column(col1),
                sg.Text(' ' * 5), sg.Button(image_filename='cam5_100x96.png', image_size=(100,90), border_width=(0), key='snip')],
                ]

    window = sg.Window('QuickSnip', layout, no_titlebar=True, grab_anywhere=True, keep_on_top=True, resizable=True)

    while True:
        event, values = window.read()
        if event == 'x' or event is None:
            break
        if event == 'snip':
            try:
                print('Saving screenshot...please wait')
                window.Refresh()
                d = d3dshot.create()
                vF = values['-vFilename-']
                vP = values['-vFolder-']
                d.screenshot_to_disk(directory=vP, file_name=str(vF + '.png'))
                print(' ')
            except:
                print('ERROR: Enter filename and folder path')
                window.Refresh()
                window.read(timeout=5000)
                print(' ')

if __name__ == '__main__':
    Launcher()
            