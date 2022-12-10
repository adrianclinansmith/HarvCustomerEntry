"""
How to Create an Excel Data Entry Form in 10 Minutes Using Python:
https://www.youtube.com/watch?v=svcv8uub0D0

To make this a standalone executable:
- $ pyinstaller --onefile --noconsole myScript.py
- The executable is dist\myScript.py
"""

import pandas as pd
import PySimpleGUI as gui

EXCEL_FILE: str = ""
df: pd.DataFrame = None
headings = ["Name", "Country"]

# Theme & Layout

gui.theme("darkteal9")
layout = [
    [gui.Text("File", key="FileText", size=15),
        gui.Input(key="File", enable_events=True, visible=False), gui.FileBrowse()],
    [gui.Text("Name", size=15), gui.Input(key="Name", enable_events=True)],
    [gui.Text("Country", size=15), gui.Input(key="Country", enable_events=True)],
    [gui.Submit()],
    [gui.Table([], headings=headings, key="Table", enable_click_events=False, enable_events=True)],
    [gui.Button(button_text="Delete", disabled=True, key="Delete")]
]

# Popup Constructor

def yesNoPopup(text: str):
    layout = [[gui.Text(text)], [gui.Yes(size=10), gui.No(size=10)]]
    return gui.Window("", layout, disable_close=True).read(close=True)[0]

def writeErrorPopup():
    message = "The data could not be updated.\n\n"
    message += f"Please ensure {EXCEL_FILE} is closed and has write permission."
    gui.popup(message)

# Event Loop

window = gui.Window("Customer Entry", layout)
while True:
    event, values = window.read()
    print(event)
    print(values)
    print()
    if event == gui.WIN_CLOSED:
        break
    elif event == "File":
        EXCEL_FILE = values["File"]
        df = pd.read_excel(EXCEL_FILE)
        startIndex = EXCEL_FILE.rfind("/") + 1
        window["FileText"].update(EXCEL_FILE[startIndex : ])
        window["Table"].update(df.values.tolist())
    elif event == "Submit":
        if df is None:
            gui.popup("You must select a file.")
            continue
        yesNoEvent = yesNoPopup(f"Save this entry to {EXCEL_FILE} ?")
        if yesNoEvent != "Yes":
            continue
        try:
            entries = {"Name": values["Name"], "Country": values["Country"]}
            newDf = df.append(entries, ignore_index=True)
            newDf.to_excel(EXCEL_FILE, index=False)
            window["Table"].update(newDf.values.tolist())
            df = newDf
        except:
            writeErrorPopup()
    elif event == "Table":
        window["Delete"].update(disabled=bool(values["Table"]))
    elif event == "Delete":
        yesNoEvent = yesNoPopup("Delete selected entries?")
        if yesNoEvent != "Yes":
            continue
        try:
            newDf = df.drop(index=window["Table"])
            newDf.to_excel(EXCEL_FILE, index=False)
            window["Table"].update(df.values.tolist())
            df = newDf
        except:
            writeErrorPopup()

window.close()