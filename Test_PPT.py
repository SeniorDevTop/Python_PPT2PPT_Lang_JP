from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
import PPTManager

# Define the function that will be called when the button is clicked
def onFileOpen():
    text = edit.get('1.0', 'end-1c')
    # Do something with the text...
    path = filedialog.askopenfilename()
    edit.delete('1.0', 'end')
    edit.insert('1.0', path)

def onConvert():
    strPath = edit.get("1.0", "end-1c")
    strLang = combo.get()
    if strLang == '' or strPath == '':
        return
    if strLang == "English":
        strLang = "en"
    else:
        strLang = "ko"

    PPTManager.main(strPath, strLang)

# Create the main window
root = tk.Tk()
root.title('PPT Manager')
root.configure(background='black')

# Create the top frame for the edit control and button
top_frame = tk.Frame(root, bg='black')
top_frame.pack(side='top', padx=10, pady=10)

# Create the edit control
edit = tk.Text(top_frame, height=1, width=50, fg='white', bg='black')
edit.pack(side='left')

# Create the button
button = tk.Button(top_frame, text='Open', command=onFileOpen, bg='white', fg='black', background='#FFC107', foreground='white')
button.pack(side='left', padx=10)

# Create the bottom frame for the combobox and button
bottom_frame = tk.Frame(root, bg='black')
bottom_frame.pack(side='bottom', padx=10, pady=(0, 20))

# Create the combobox
options = ['Korean', 'English']
combo = tk.ttk.Combobox(bottom_frame, values=options, state='readonly', background='#FFC107', foreground='white')
combo.pack(side='left', padx=(10, 20))

# Create the conversion button
convert_button = tk.Button(bottom_frame, text='Convert', bg='white', fg='black', background='#FFC107', foreground='white', command=onConvert)
convert_button.pack(side='left', padx=10)

# Start the main loop
root.mainloop()

