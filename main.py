import ctypes
import os
import sys

import PySimpleGUI as sg
from win32com.client import Dispatch

allStartupPrograms = []


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def check_close(event):
    if event == 'Close' or event is None:
        sys.exit(0)


def remove_startup():
    layout = [[sg.Text('Filename')],
              [sg.Button("Remove"), sg.Button('Back')],
              [sg.Listbox(allStartupPrograms, size=(30, 6), enable_events=True, key='_LIST_')]]

    window = sg.Window('Remove From Startup', layout)

    while True:
        event, values = window.read()
        if event == 'Remove':
            file = values["_LIST_"][0]
            path = os.path.join(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp", str(file) + ".lnk")
            os.remove(path)
            allStartupPrograms.remove(file)
            window.Element('_LIST_').Update(allStartupPrograms)

        elif event == 'Back' or event is None:
            window.close()
            break
        check_close(event)


def check_null(values):
    if values[0] == '':
        sg.popup('Please select a file')
        return False
    else:
        return True


def save_startup(event, values):
    if event == "OK":
        FileName = (os.path.basename(values[0]))

        path = os.path.join(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp", FileName + ".lnk")
        target = values[0]
        wDir = values[0]
        icon = values[0]
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wDir
        shortcut.IconLocation = icon
        shortcut.save()
        sg.popup('File added to startup')


if is_admin():
    # Code of your program here
    sg.theme('Dark Blue 3')  # please make your creations colorful

    layout = [[sg.Text('Filename')],
              [sg.Input(), sg.FileBrowse()],
              [sg.OK(), sg.Button('Close'), sg.Button('Remove')]]

    window = sg.Window('Select file for startup', layout)

    done = False
    while not done:
        event, values = window.read()
        if event == "Remove":
            allStartupPrograms = []
            for file in os.listdir(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"):
                if file.endswith(".lnk"):
                    allStartupPrograms.append(file.removesuffix('.lnk'))
            remove_startup()
        elif event == "OK":
            if not check_null(values):
                continue
            save_startup(event, values)
        else:
            check_close(event)

else:
    # Re-run the program with admin rights
    sg.popup('Please start in admin')

    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
