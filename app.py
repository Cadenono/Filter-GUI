import tkinter as tk
import tkinter.filedialog as fdialog
import pandas as pd
from glob import glob
import xlrd
import os


WINDOW_TITLE = "Filter App"
WINDOW_WIDTH = 700
WINDOW_HEIGHT = 800
my_photo = 'rsz_panda.png'
x = False
filepath = ''
entry_box_input = ''
check = False 


class FilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.canvas = tk.Canvas(self.root, height=WINDOW_HEIGHT, width=WINDOW_WIDTH)
        self.create_background(my_photo)
        self.canvas.pack()
        self.top_frame = tk.Frame(self.root)
        self.top_frame.place(relx=0.5, rely=0.1,
                             relwidth=0.75, relheight=0.55, anchor='n')
        self.entry_box()
        self.create_buttons()

    def create_background(self, photo):
        bg_img = tk.PhotoImage(file=photo)
        self.canvas.background = bg_img
        bg_label = tk.Label(self.canvas, image=bg_img)
        bg_label.place(relwidth=1, relheight=1)

    def open_file(self):
        filepaths = fdialog.askdirectory(
            parent=self.root, title='Choose folder')
        open_label = tk.Label(
            self.top_frame, text="Successful upload. Enter your vehicle number.")
        open_label.pack()
        global x
        x = True
        global filepath
        filepath = filepaths


    def entry_box(self):
        global entry_box_input
        entry = tk.Entry(self.canvas, font=30)
        entry.insert(0, "Enter Vehicle No")
        entry.place(relwidth=0.44, relheight=0.1, rely=0.7, relx=0.35, anchor='n')
        entry_box_input = entry

    def retrieve_entry_box_input(self):
        global entry_box_input
        global x
        global filepath
        global check 
        if x:
            try:
                filenames = [os.path.abspath(files) for files in glob(filepath + "\*.xlsm")]
                dfs = [pd.read_excel(file, sheet_name=2, header=0) for file in filenames]
                if len(dfs) == 1: 
                    merged_df = df 
                else: 
                    
                    merged_df =pd.concat(dfs)
                
            except: 
                text = tk.Text(self.top_frame)
                text.insert(tk.END, 'Please try again. Make sure to close your file before clicking "Filter".')
                text.pack(side=tk.BOTTOM)
            finally: 
                if entry_box_input.get() not in set(merged_df['Veh No']):
                    text = tk.Text(self.top_frame)
                    text.insert(tk.END, 'Error. No such vehicle. Reopen GUI to try again.')
                    text.pack(side=tk.BOTTOM)
                    

          
                else: 
                    filtered_file = merged_df[merged_df['Veh No'] == entry_box_input.get()]
                    num_rows = filtered_file.shape[0]
                    text = tk.Text(self.top_frame)
                    text.insert(tk.END,
                                str(filtered_file) + "\n" + "Filtered. There are {} occurrences of vehicle {}.".format(
                                    num_rows, entry_box_input.get()))
                    text.pack(side=tk.BOTTOM)

                    filtered_file.to_csv(filepath + "\ " + entry_box_input.get() + " .csv", index=False, header=True)
                  
                
        else: 
            text = tk.Text(self.top_frame)
            text.insert(tk.END, 'Upload a file.')
            text.pack(side=tk.BOTTOM)
           

    def create_buttons(self): 
        filter_button = tk.Button(self.canvas, text="Filter", font="none 12 bold", command=self.retrieve_entry_box_input)
        filter_button.place(relx=0.75, rely=0.7, relheight=0.1, relwidth=0.25, anchor='n')
        upload_button = tk.Button(self.top_frame, text='Upload File', command=self.open_file)
        upload_button.pack(side=tk.TOP)


root = tk.Tk()
app = FilterApp(root)
root.mainloop()

