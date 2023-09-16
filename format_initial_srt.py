import PySimpleGUI as sg
import srt

layout = [
    [sg.Text("Input SRT file:"), sg.In(), sg.FileBrowse(key="-BROWSE_INPUT")],
    [sg.Text("Output SRT file:"), sg.In(), sg.FileSaveAs(key="-BROWSE_OUTPUT")],
    [sg.Button("Format initial SRT file",key="-INITIAL_FORMAT"),sg.Button("Create script for voice-over")]
]

window = sg.Window("Subtitle Tools for The Week",layout)

while True:
    event, values = window.read()

    if event == "-INITIAL_FORMAT":
        file_path = layout[0][1].get()
        target_file_path = layout[1][1].get()

        srt_file = open(file_path, 'r')
        target_file = open(target_file_path,'w')
        modified_srt_file = srt_file
        subs = list(srt.parse(srt_file))

        for idx, sub in enumerate(subs):
            if sub.content.count('\n') == 0:
                sub.content = sub.content + '\n'

        target_file.write(srt.compose(subs))

        srt_file.close()
        target_file.close()

        sg.popup_ok("Formatted SRT successfully saved",title="Success")

    # End program if user closes window or
    # presses the OK button
    if event == "OK" or event == sg.WIN_CLOSED:
        break

window.close()




