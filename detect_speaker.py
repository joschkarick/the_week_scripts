import PySimpleGUI as sg
import os
import sys
import csv
import datetime
from pyannote.audio.pipelines.utils.hook import ProgressHook

def popup_error(message, title=None):
    layout = [
        [sg.Text(message,background_color=main_background_color,font=main_font)],
        [sg.Button('Retry',background_color=second_background_color,font=main_font)]
    ]

    window = sg.Window(title if title else message, layout)
    event, values = window.read(close=True)
    return

def detect_speaker():
    from pyannote.audio import Pipeline

    pipeline = Pipeline.from_pretrained(
        "pyannote/speaker-diarization-3.0",
        use_auth_token="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")

    # send pipeline to GPU (when available)
    #import torch
    #pipeline.to(torch.device("cuda"))

    # apply pretrained pipeline
    with ProgressHook() as hook:
        diarization = pipeline(layout[0][1].get(), hook=hook)

    with open(layout[1][0].get(), 'w', newline='', encoding='UTF8') as f:
        writer = csv.writer(f)

        header = ['Start', 'End', 'Speaker', 'Start in s', 'End in s', 'Duration in s']

        writer.writerow(header)

        # print the result
        for turn, _, speaker in diarization.itertracks(yield_label=True):
            data = [str(datetime.timedelta(seconds=turn.start)),str(datetime.timedelta(seconds=turn.end)),speaker,turn.start,turn.end,turn.end-turn.start]
            writer.writerow(data)

    sg.popup_ok("Speakers successfully extracted",title="Success",background_color=main_background_color, font=main_font)

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
        sg.Text("Input WAV file:",background_color=main_background_color, font=main_font),
        sg.In(font=main_font), sg.FileBrowse(key="-BROWSE_INPUT", button_color=second_background_color, font=main_font)
    ],
    [
        sg.In(key="-DETECT_SPEAKER",visible=False,enable_events=True),
        sg.FileSaveAs("Speaker recognition",default_extension=".csv",file_types=(("CSV file (csv)",".csv"),),button_color=second_background_color, font=main_font),
    ]
]

window = sg.Window("Tools for The Week",layout,icon=resource_path("theweek.ico"), background_color=main_background_color) 

while True:
    event, values = window.read()

    if event == "-DETECT_SPEAKER":
        detect_speaker()

    # End program if user closes window
    if event == sg.WIN_CLOSED:
        break

window.close()

