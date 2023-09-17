import PySimpleGUI as sg
import srt
from docx import Document
from docx.shared import Cm


def popup_error(message, title=None):
    layout = [
        [sg.Text(message)],
        [sg.Button('Retry')]
    ]

    window = sg.Window(title if title else message, layout)
    event, values = window.read(close=True)
    return

def ask_for_file_target(message, title=None, button_text='Create'):
    layout = [
        [sg.Text(message)],
        [sg.Input(key='-TARGET_FILE'), sg.FileBrowse(key='-BROWSE_TARGET_FILE')],
        [sg.Button(button_text,key='CREATE'), sg.Button('Cancel',key='-CANCEL')]
    ]

    window = sg.Window(title if title else message, layout)
    event, values = window.read(close=True)
    return values['-TARGET_FILE'] if event == 'CREATE' else None

def create_formatted_srt():
    file_path = layout[0][1].get()
    srt_file = open(file_path, 'r')
    subs = list(srt.parse(srt_file))

    target_file_path = layout[1][1].get()
    target_file = open(target_file_path,'w')

    for idx, sub in enumerate(subs):
        if sub.content.count('\n') == 0:
            sub.content = sub.content + '\n'

    target_file.write(srt.compose(subs))

    srt_file.close()
    target_file.close()

    sg.popup_ok("Formatted SRT successfully saved",title="Success")

def create_script_for_voice_over():
    file_path = layout[0][1].get()
    srt_file = open(file_path, 'r')
    subs = list(srt.parse(srt_file))

    target_file_path = layout[2][1].get()
    #target_file_path = ask_for_file_target(message="Please choose target .docx file", button_text='Create script',title='Create script for voice-over')

    script = Document()

    for sub in subs:
        script.add_heading('[INSERT NARRATOR/SPEAKER HERE]', level=1)
        script.add_paragraph(sub.content)

    try:
        script.save(target_file_path)

        sg.popup_ok("Formatted SRT successfully saved",title="Success")
        return
    except PermissionError:
        popup_error("Could not save file. Is the target file still open in Word?",title="Error")

layout = [
    [
        sg.Text("Input SRT file:"),
        sg.In(), sg.FileBrowse(key="-BROWSE_INPUT")
    ],
    [
        sg.Text("Output SRT file:"),
        sg.In(), sg.FileSaveAs(key="-BROWSE_OUTPUT"),
        sg.Button("Format initial SRT file",key="-CREATE_FORMATTED_SRT")
    ],
    [
        sg.Text("Output script for voice-over"),
        sg.In(), sg.FileSaveAs(key="-BROWSE_OUTPUT_SCRIPT"),
        sg.Button("Create script for voice-over",key="-CREATE_WORD_VOICE_OVER")
    ]
]

window = sg.Window("Subtitle Tools for The Week",layout)

while True:
    event, values = window.read()

    if event == "-CREATE_FORMATTED_SRT":
        create_formatted_srt()

    if event == "-CREATE_WORD_VOICE_OVER":
        create_script_for_voice_over()

    # End program if user closes window or
    # presses the OK button
    if event == "OK" or event == sg.WIN_CLOSED:
        break

window.close()

