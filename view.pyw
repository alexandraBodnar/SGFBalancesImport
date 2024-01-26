import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
from openpyxl.utils.exceptions import InvalidFileException
import unionbalances
import traceback

PADL = 30  # padding left - to align items

# global variables to save paths taken from view
savingspath = ''
grantpath = ''
alumnipath = ''


class Application(tk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.add_widgets()

    def add_widgets(self):
        self.savingslabel = tk.Label(root, text='Select savings account file:')
        self.savingslabel.place(x=PADL, y=30)
        self.savingsbutton = ttk.Button(root, text='Browse', command=self.open_savings)
        self.savingsbutton.place(x=PADL, y=55)
        self.savingspathlabel = tk.Label(root, text='')
        self.savingspathlabel.place(x=PADL + self.savingsbutton.winfo_width() + 80, y=55)

        self.grantlabel = tk.Label(root, text='Select grant account file:')
        self.grantlabel.place(x=PADL, y=105)
        self.grantbutton = ttk.Button(root, text='Browse', command=self.open_grant)
        self.grantbutton.place(x=PADL, y=130)
        self.grantpathlabel = tk.Label(root, text='')
        self.grantpathlabel.place(x=PADL + self.savingsbutton.winfo_width() + 80, y=130)

        self.alumnilabel = tk.Label(root, text='Select ring-fenced account file:')
        self.alumnilabel.place(x=PADL, y=180)
        self.alumnibutton = ttk.Button(root, text='Browse', command=self.open_alumni)
        self.alumnibutton.place(x=PADL, y=205)
        self.alumnipathlabel = tk.Label(root, text='')
        self.alumnipathlabel.place(x=PADL + self.savingsbutton.winfo_width() + 80, y=205)

        self.statuslabel = tk.Label(root, text='')
        self.statuslabel.place(relx=0.5, y=255, anchor='center')

        self.submitbutton = ttk.Button(root, text='Import', command=self.import_with_path)
        self.submitbutton.pack(side='bottom', pady=20)

    # general function to open the relevant file - identified by number
    # necessary so that the path gets saved in the relevant variable
    # 1 - savings
    # 2 - grant
    # 3 - ringfenced
    def open_file(self, r):
        global savingspath, grantpath, alumnipath
        file = filedialog.askopenfile(mode='r')
        if file:
            filepath = os.path.abspath(file.name)
            if r == 1:
                self.savingspathlabel.config(text=str(filepath))
                savingspath = filepath
            elif r == 2:
                self.grantpathlabel.config(text=str(filepath))
                grantpath = filepath
            else:
                self.alumnipathlabel.config(text=str(filepath))
                alumnipath = filepath

    def open_savings(self):
        self.open_file(1)

    def open_grant(self):
        self.open_file(2)

    def open_alumni(self):
        self.open_file(3)

    # call import from unionbalances.py
    # if successful, inform user that the file is in downloads folder
    # if not, inform them of the problem
    def import_with_path(self):
        try:
            unionbalances.importAll(savingspath, grantpath, alumnipath)
            self.statuslabel.config(text="Done! File saved in your Downloads folder")
        except PermissionError:
            self.statuslabel.config(text="Please close the import template file")
        except InvalidFileException:
            self.statuslabel.config(text="One of the files is invalid, please double check")
        except:
            traceback.print_exc()
            self.statuslabel.config(text="Unknown error")


root = tk.Tk()
root.title("Balances import app")
root.geometry('500x350')
app = Application(master=root)
app.mainloop()
