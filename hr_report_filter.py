from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from datetime import datetime
import pandas as pd
import time
import xlsxwriter

# global variables
hr_global_filename = ""
output_folder = ""

# functions
# define run function that is the backend data cleaning/filtering
def run():
    
#hr_file = hr_global_filename
    
    df = pd.read_excel(hr_global_filename)
        
    tp_only = df[((df['Global Employee Subgroup'] == 'Associate') & (df['Level'] == 'Talent Pool 1')) | (df['Global Employee Subgroup'] == 'SL1') & (df['Level'] == 'Talent Pool 2')]
        
    tp_final = tp_only.drop_duplicates(subset ='Last Name').copy()
        
    tp_final = tp_final[['First Name', 'Last Name','Functional Area (Person)','Organizational Area','Email','Work Location','Country','Level','Section Custom Field Value']]
        
    # Define today's date to be used for output file name
    todaysdate = datetime.today().strftime('%m-%d-%Y')
    # Define filename of output excel file
    path = output_folder
    name = f"HR Global_Talent_Pool_{todaysdate}.xlsx"
    outpath = path + "/" + name
    # Define method to export dataframe
    writer = pd.ExcelWriter(outpath, engine='xlsxwriter')
    # Write dataframe to shee
    tp_final.to_excel(writer, sheet_name ='HR Global Talent Pool',index=False)
    # Create xlsxwriter objects for formatting
    workbook_object= writer.book
    worksheet_object = writer.sheets['HR Global Talent Pool']
    worksheet_object.set_column('A:A', 17) # First Name
    worksheet_object.set_column('B:B', 18.5) # Last Name
    worksheet_object.set_column('C:C', 32) # Functional Area (Person)
    worksheet_object.set_column('D:D', 17) # Organizational Area
    worksheet_object.set_column('E:E', 52) # Email
    worksheet_object.set_column('F:F', 13) # Work Location
    worksheet_object.set_column('G:G', 11.5) # Country
    worksheet_object.set_column('H:H', 11) # Level
    worksheet_object.set_column('I:I', 255) # Comments

    writer.save()

# define function that is linked to "import" button in GUI and will allow user to search for file to import
def hr_open():
    global hr_global_filename
    hr_global_filename = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    hr_data = Label(root, text=hr_global_filename, width=40, borderwidth=2, relief="groove")
    hr_data.grid(row=0, column=1)

def find_folder():
    global output_folder
    output_folder = filedialog.askdirectory(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a folder",
    )
    folder = Label(root, text=output_folder, width=40, borderwidth=2, relief="groove")
    folder.grid(row=1, column=1)

# Set up GUI window
root = Tk()
root.geometry("515x140")
root.wm_title("Talent Pool Data Filter")

# labels
hr_label = Label(root, text="HR Global File")
hr_label.grid(row=0, column=0, padx=10, pady=10)

hr_blank = Label(root, width=40, borderwidth=2, relief="groove")
hr_blank.grid(row=0, column=1)

source_label = Label(root, text="Output Folder")
source_label.grid(row=1, column=0, padx=10, pady=10)

source_blank = Label(root, width=40, borderwidth=2, relief="groove")
source_blank.grid(row=1, column=1, padx=10, pady=10)

# buttons
run = Button(root, text="Run", width=10, bg="gray60", fg="black", command=run)
run.grid(row=4, column=2, padx=10, pady=10)

import1 = Button(root, text="Import", width=10, command=hr_open)
import1.grid(row=0, column=2, padx=10, pady=10)

find = Button(root, text="Find", width=10, command=find_folder)
find.grid(row=1, column=2, padx=10, pady=10)


root.mainloop()