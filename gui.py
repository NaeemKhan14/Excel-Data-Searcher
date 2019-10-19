from tkinter import *
import tkinter as tk
from tkinter import filedialog
from processor import Processor


class GUI(Tk):

    def __init__(self):
        global user_input, filename_str, filename, search_col, result_col, finish_label

        super(GUI, self).__init__()
        self.title("Excel Searcher")
        self.minsize(640, 100)

        # Input parameter label and Entry settings
        tk.Label(self, text="Input parameters to search for").grid(row=0)

        user_input = tk.Entry(self, width=100)
        user_input.grid(row=0, column=1)
        user_input.focus_set()

        # File selection label
        tk.Label(self, text="Input file").grid(row=1)
        # File Entry's text which will change with the file name selected from browse button
        filename_str = tk.StringVar()
        filename = tk.Entry(self, textvariable=filename_str, width=100)
        filename.grid(row=1, column=1)

        # Browse button to look for the file on the computer
        tk.Button(self, text="Browse", command=self.load_file).grid(row=1, column=2)

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

    # Get values of all the Entry boxes on this button press
    def start_button(self):
        processor = Processor(str(filename.get()), str(user_input.get()))
        processor.process_rows(search_col=int(search_col.get()), result_col=int(result_col.get()))
        finish_label.config(text="Finished Processing")
