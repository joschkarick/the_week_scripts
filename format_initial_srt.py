import tkinter as tk
import srt
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
target_file_path = filedialog.asksaveasfilename()

srt_file = open(file_path, 'r')
target_file = open(target_file_path,'w')
modified_srt_file = srt_file
subs = list(srt.parse(srt_file))

for idx, sub in enumerate(subs):
    if sub.content.count('\n') == 0:
        sub.content = sub.content + '\n'
        print('Modifying line ', idx)
    print(sub)

target_file.write(srt.compose(subs))

srt_file.close()
target_file.close()


