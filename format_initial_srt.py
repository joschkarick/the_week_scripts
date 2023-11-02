import PySimpleGUI as sg
import srt
import os
import sys
#import openpyxl
import csv
#from openpyxl import load_workbook
#from openpyxl.worksheet.table import Table
from docx import Document
from docx.shared import Cm


def popup_error(message, title=None):
    layout = [
        [sg.Text(message,background_color=main_background_color,font=main_font)],
        [sg.Button('Retry',background_color=second_background_color,font=main_font)]
    ]

    window = sg.Window(title if title else message, layout)
    event, values = window.read(close=True)
    return

def create_formatted_srt():
    srt_file = load_srt()
    subs = list(srt.parse(srt_file))

    srt_file = load_srt()
    subs = list(srt.parse(srt_file))

    for idx, sub in enumerate(subs):
        if sub.content.count('\n') == 0:
            sub.content = sub.content + '\n'

    target_file_path = layout[2][0].get()

    target_file = open(target_file_path,'w', encoding='utf-8')
    target_file.write(srt.compose(subs))

    srt_file.close()
    target_file.close()

    sg.popup_ok("Formatted SRT successfully saved",title="Success",background_color=main_background_color, font=main_font)

def create_script_for_voice_over():
    srt_file = load_srt()
    subs = list(srt.parse(srt_file))

    target_file_path = layout[3][0].get()

    script = Document()

    previous_speaker = ''

    speaker_timecodes = load_speaker_timecode_csv()
    speaker_timecodes_keys = [key for key in speaker_timecodes.keys()]
    speaker_timecodes_as_float = [float(key) for key in speaker_timecodes_keys[1:]]

    for sub in subs:
        closest_speaker_timecode_as_float = closest(speaker_timecodes_as_float, sub.start.total_seconds())
        speaker = speaker_timecodes[str(closest_speaker_timecode_as_float)]

        if previous_speaker != speaker:
            script.add_heading(speaker, level=1)
            previous_speaker = speaker

        script.add_paragraph(sub.content.replace("\n"," "))

    try:
        script.save(target_file_path)

        sg.popup_ok("Script DOCX successfully saved",title="Success",background_color=main_background_color, font=main_font)
        return
    except PermissionError:
        popup_error("Could not save file. Is the target file still open in Word?",title="Error")

def closest(lst, target):
    return lst[min(range(len(lst)), key = lambda i: abs(lst[i]-target))]

#def add_srt_to_xlsx():
#    target_xlsx_path = layout[4][0].get()
#    wb = openpyxl.load_workbook(target_xlsx_path)
#    ws = wb['Subs DE EP 1']
#    print(ws.title)
#
#    # TODO: Identify correct column
#
#    cell = ws["B1:B6000"]
#    print(cell)

def load_srt():
    try:
        file_path = layout[0][1].get()
        srt_file = open(file_path, 'r', encoding='utf-8')

        return srt_file
    except FileNotFoundError:
        popup_error("SRT input file was not found. Is the path correct?",title="Error")

def load_speaker_timecode_csv():
    try:
        file_path = layout[1][1].get()
        speaker_timecode_csv_file = open(file_path, 'r', encoding='utf-8')

        with open(file_path, mode='r') as csv_file:
            reader = csv.reader(csv_file, delimiter=';', quotechar='"')

            speaker_timecodes = {rows[3]:rows[2] for rows in reader}

        return speaker_timecodes
    except FileNotFoundError:
        popup_error("Speaker Timecode CSV input file was not found. Is the path correct?",title="Error")
    except IndexError:
        popup_error("Speaker Timecode CSV could not be read. Make sure to use a semicolon (;) as a separator and quotation marks for text.")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

main_background_color = "#434243"
second_background_color = "#faa61a"
main_font_family = "Lato"
main_font = (main_font_family,"12")

layout = [
    [
        sg.Text("Input Subtitle SRT file:",background_color=main_background_color, font=main_font),
        sg.In(font=main_font), sg.FileBrowse(key="-BROWSE_INPUT", button_color=second_background_color, font=main_font)
    ],
    [
        sg.Text("Input Speaker Timecodes CSV file:",background_color=main_background_color, font=main_font),
        sg.In(font=main_font), sg.FileBrowse(key="-BROWSE_INPUT", button_color=second_background_color, font=main_font)
    ],
    [
        sg.In(key="-SRT_TARGET_FILLED",visible=False,enable_events=True),
        sg.FileSaveAs("Save SRT file as formatted SRT file",default_extension=".srt",file_types=(("SRT files",".srt"),),button_color=second_background_color, font=main_font),
    ],
    [
        sg.In(key="-SCRIPT_TARGET_FILLED",visible=False,enable_events=True),
        sg.FileSaveAs("Merge SRT file and speaker timecodes into a script",default_extension=".docx",file_types=(("Word files (docx)",".docx"),),button_color=second_background_color, font=main_font),
    ],
#    [
#        sg.In(key="-XLSX_ADD_SRT",visible=False,enable_events=True),
#        sg.FileSaveAs("Add SRT to XLSX",default_extension=".xlsx",file_types=(("Excel files (xlsx)",".xlsx"),),button_color=second_background_color, font=main_font),
#    ]
]

window = sg.Window("Subtitle Tools for The Week",layout,icon=resource_path("theweek.ico"), background_color=main_background_color) 

while True:
    event, values = window.read()

    if event == "-SRT_TARGET_FILLED":
        create_formatted_srt()

    if event == "-SCRIPT_TARGET_FILLED":
        create_script_for_voice_over()

#    if event == "-XLSX_ADD_SRT":
#        add_srt_to_xlsx()
        
    # End program if user closes window
    if event == sg.WIN_CLOSED:
        break

window.close()

