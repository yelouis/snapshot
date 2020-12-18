# img_viewer.py
import PySimpleGUI as sg
import os.path
import docx2txt
import re
import json
import csv

def scannerFunction(documentPath, csvPath):
    #----------01_Import File Name----------
    document = docx2txt.process(documentPath)  #Change filename here
    with open(csvPath, newline = '') as f:
        reader = csv.reader(f)
        keywords = [row[0] for row in reader]

    #02_-----------Declare Variables-----------
    foundKeywords = []
    #-----------03_Extract Elements From the Word File-----------
    # for para in document.paragraphs:
    #     text = para.text
    #     for keyword in keywords:
    #         if keyword.isupper():
    #             keyword_list = re.findall(r'(%s)' % keyword, text)
    #         else:
    #             keyword_list = re.findall(r'(%s)' % keyword, text, re.IGNORECASE)
    #         if len(keyword_list) != 0:
    #             for i in range(len(keyword_list)):
    #                 foundKeywords.append(keyword)

    # tbl = list(document.tables)
    # for table in tbl:
    #     for rw in table.rows:
    #         for i in range(len(rw.cells)):
    #             text = rw.cells[i].text
    for keyword in keywords:
        if keyword.isupper():
            keyword_list = re.findall(r'(%s)' % keyword, document)
        else:
            keyword_list = re.findall(r'(%s)' % keyword, document, re.IGNORECASE)
        if len(keyword_list) != 0:
            for i in range(len(keyword_list)):
                foundKeywords.append(keyword)

    #-----------04_Create Output-----------
    foundKeywordsFreq = {}
    for item in foundKeywords:
        if (item in foundKeywordsFreq):
            foundKeywordsFreq[item] += 1
        else:
            foundKeywordsFreq[item] = 1

    foundKeywordsFreq = list(foundKeywordsFreq.items())

    output = []
    for foundKeyword in foundKeywordsFreq:
        output.append(str(foundKeyword[0]) + ": " + str(foundKeyword[1]))
    return output

file_list_column = [
    [
        sg.Text("Word File"),
        sg.In(size=(25, 1), enable_events=True, key="-WORD-"),
        sg.FileBrowse()
    ],
    [
        sg.Text("CSV File"),
        sg.In(size=(25, 1), enable_events=True, key="-CSV-"),
        sg.FileBrowse()
    ],
    [
        sg.Button("OK", enable_events=True, key="-OK-")
    ]
]

image_viewer_column = [
    [sg.Text("Snapshot of keywords found in document:")],
    [sg.Text(size=(40, 1), key="-TOUT-")],
    [sg.Listbox(values=[], enable_events=True, size=(40, 20), key="-WORD LIST-")]
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(image_viewer_column)
    ]
]

window = sg.Window("Snapshot", layout)

# Run the Event Loop
while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    if event == "-OK-":
        if len(values["-WORD-"]) > 0 and len(values["-CSV-"]) > 0:
            try:
                output = scannerFunction(values["-WORD-"], values["-CSV-"])
                window["-TOUT-"].update("Completed")
                window["-WORD LIST-"].update(output)
            except:
                window["-TOUT-"].update("Error Occured when analyzing files")
                window["-WORD LIST-"].update([])

window.close()
