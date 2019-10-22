from tkinter import *
from tkinter.messagebox import showinfo, showwarning
import tkinter as tk
from tkinter import filedialog, scrolledtext
import os.path
from processor import Processor


class GUI(Tk):

    def __init__(self):
        global user_input, regex_input, filename_str, filename, search_col, result_col, finish_label

        super(GUI, self).__init__()
        self.title("Excel Searcher")
        self.minsize(640, 100)

        # Input parameter label and Textbox settings
        tk.Label(self, text="Input parameters to search for").grid(row=0)

        user_input = scrolledtext.ScrolledText(self, height=5, width=73, wrap=WORD)
        user_input.bind("<Return>", self.focus_next_window)
        user_input.bind("<Tab>", self.focus_next_window)
        user_input.grid(row=0, column=1)
        user_input.focus_set()

        # Regex parameter label and Textbox settings
        tk.Label(self, text="Regex parameters to search for").grid(row=1)

        regex_input = scrolledtext.ScrolledText(self, height=5, width=73, wrap=WORD)
        regex_input.bind("<Return>", self.focus_next_window)
        regex_input.bind("<Tab>", self.focus_next_window)
        regex_input.grid(row=1, column=1)

        # File selection label
        tk.Label(self, text="Input file").grid(row=2)
        # File Entry's text which will change with the file name selected from browse button
        filename_str = tk.StringVar()
        filename = tk.Entry(self, textvariable=filename_str, width=100)
        filename.grid(row=2, column=1)

        # Browse button to look for the file on the computer
        tk.Button(self, text="Browse", command=self.load_file).grid(row=2, column=2)

        # Label for searching column and search column settings
        tk.Label(self, text="Column to search in").grid(row=3)
        search_col = tk.Entry(self, width=100)
        search_col.grid(row=3, column=1)

        # Label and Entry settings for result column
        tk.Label(self, text="Column to get results from").grid(row=4)
        result_col = tk.Entry(self, width=100)
        result_col.grid(row=4, column=1)

        # Submit button
        tk.Button(self, text="Start", command=self.start_button).grid(row=5, column=1)

        # Finished Label
        finish_label = tk.Label(self, text="")
        finish_label.grid(row=6, column=1)

    # Browse dialog to select the file from computer
    def load_file(self):
        filename_str.set(filedialog.askopenfilename(master=self))

    # Makes the textbox go to next widget when pressing tab or enter key
    def focus_next_window(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    # Returns True if the given textbox is empty
    def textbox_isEmpty(self, textbox):
        if len(textbox.get('0.0', 'end')) <= 1:
            return True

        return False

    # Returns True if the given Entry is empty
    def entry_isEmpty(self, entrybox):
        if len(entrybox.get()) == 0:
            return True

        return False

    # Get values of all the Entry boxes on this button press
    def start_button(self):
        if self.textbox_isEmpty(user_input) and self.textbox_isEmpty(regex_input):
            showinfo("No input given", "Please write something in one of the input boxes")
        elif self.entry_isEmpty(filename):
            showinfo("No file given", "Please select a file to continue")
        elif self.entry_isEmpty(search_col):
            showinfo("Search column empty", "Please write a value in search box")
        elif self.entry_isEmpty(result_col):
            showinfo("Result column empty", "Please write a value in result box")
        elif not self.entry_isEmpty(filename) and not os.path.isfile(str(filename.get())):
            showwarning("File does not exist", "Please select a valid file")
        else:
            processor = Processor(str(filename.get()), user_input.get('1.0', 'end-1c'), regex_input.get('1.0', 'end-1c'))
            processor.process_rows(search_col=int(search_col.get()), result_col=int(result_col.get()))
            finish_label.config(text="Finished Processing")
