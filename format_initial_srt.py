import PySimpleGUI as sg
import srt
import os
import sys
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

    target_file_path = layout[1][0].get()

    target_file = open(target_file_path,'w', encoding='utf-8')
    target_file.write(srt.compose(subs))

    srt_file.close()
    target_file.close()

    sg.popup_ok("Formatted SRT successfully saved",title="Success",background_color=main_background_color, font=main_font)

def create_script_for_voice_over():
    srt_file = load_srt()
    subs = list(srt.parse(srt_file))

    target_file_path = layout[2][0].get()

    print(target_file_path)

    script = Document()

    for sub in subs:
        script.add_heading('[Placeholder for narrator/speaker]', level=1)
        script.add_paragraph(sub.content.replace("\n",""))

    try:
        script.save(target_file_path)

        sg.popup_ok("Script DOCX successfully saved",title="Success",background_color=main_background_color, font=main_font)
        return
    except PermissionError:
        popup_error("Could not save file. Is the target file still open in Word?",title="Error")

def load_srt():
    try:
        file_path = layout[0][1].get()
        srt_file = open(file_path, 'r', encoding='utf-8')

        return srt_file
    except FileNotFoundError:
        popup_error("SRT input file was not found. Is the path correct?",title="Error")

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
        sg.Text("Input SRT file:",background_color=main_background_color, font=main_font),
        sg.In(font=main_font), sg.FileBrowse(key="-BROWSE_INPUT", button_color=second_background_color, font=main_font)
    ],
    [
        sg.In(key="-SRT_TARGET_FILLED",visible=False,enable_events=True),
        sg.FileSaveAs("Save as formatted SRT file",default_extension=".srt",file_types=(("SRT files",".srt"),),button_color=second_background_color, font=main_font),
    ],
    [
        sg.In(key="-SCRIPT_TARGET_FILLED",visible=False,enable_events=True),
        sg.FileSaveAs("Save as script for voice-over",default_extension=".docx",file_types=(("Word files (docx)",".docx"),),button_color=second_background_color, font=main_font),
    ]
]

window = sg.Window("Subtitle Tools for The Week",layout,icon=resource_path("theweek.ico"), background_color=main_background_color) 

while True:
    event, values = window.read()

    if event == "-SRT_TARGET_FILLED":
        create_formatted_srt()

    if event == "-SCRIPT_TARGET_FILLED":
        create_script_for_voice_over()

    # End program if user closes window
    if event == sg.WIN_CLOSED:
        break

window.close()

