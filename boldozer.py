# Modules for GUI
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

# Modules for HTML File Parsing
import codecs
import io
from bs4 import BeautifulSoup as bs

import re
from datetime import date

root = Tk()

def initialGui():

    # Tkinter settings
    root.title("BOLdozer 2")
    #root.geometry('300x200')
    #root.configure(bg='#3E4149')

    # (when button is clicked) get files/filepaths
    def openFile():
        global html_file_path
        html_file_path = filedialog.askopenfilename(initialdir='/', title='Open File', filetypes=(('HTML files','*.html'),('All files','*.*')))
        html_file_label = Label(root, text=f'Chosen file: {html_file_path}')
        html_file_label.grid(row=2, column=0)
        #print(html_file)

    def getDest():
        global dest_path
        dest_path = filedialog.askdirectory(initialdir='/', title='Open Directory')
        dest_path_label = Label(root, text=f'Chosen directory: {dest_path}')
        dest_path_label.grid(row=5, column=0)
        #print(dest_path)

    # Very basic user input validation
    def checkInputValidation():
        try:
            html_file_path
            dest_path
            root.quit()
        except NameError:
            messagebox.showwarning(title='Error', message="You didn\'t select anything!")

    # Widgets
    select_file = Label(root, text='Select vendorsheet.html file:', padx='50')
    browse_file = Button(root, text='Browse', highlightbackground='#3E4149', padx='50', command=openFile)
    select_dest = Label(root, text='Select destination path:', padx='50')
    directory = Button(root, text='Browse', highlightbackground='#3E4149', padx='50', command=getDest)
    flatten = Button(root, text='BOLdoze!', highlightbackground='#3E4149', padx='50', command=checkInputValidation)

    
    # Widget placement
    select_file.grid(row=0, column=0)
    browse_file.grid(row=1, column=0)
    select_dest.grid(row=3, column=0)
    directory.grid(row=4, column=0)
    flatten.grid(row=6, column=0)

    # Display window
    root.mainloop()

    return html_file_path, dest_path



def readHtmlFile(html_file_path, dest_path):
    
    # Read HTML file
    html_file_messy = codecs.open(html_file_path, 'r').read()
    html_file_stillmessy = bs(html_file_messy, 'html.parser')
    html_file = html_file_stillmessy.prettify()

    # Strip unnecessary data, create new list of needed info
    html_file_list = []
    for line in (html_file.splitlines()):
        if "<" not in line and ">" not in line:
            if "Vendor Load Sheet" not in line:
                info = line.strip()
                html_file_list.append(info)

    return html_file_list



def sortHtmlFile(html_file_list):
    
    # Define some variables
    range_of_list = range(len(html_file_list))
    order_list = []
    global load_no

    # Extract load number, which is always the 2nd item
    load_no = html_file_list[1]
    
    # Extract Total PO, Pick Stop Number
    for num in range_of_list:
        if 'Total PO(s):' in html_file_list[num]:
            total_po = int(re.findall(r'\d+',html_file_list[num]).pop())
        if 'Pick' in html_file_list[num] and '54669701' in html_file_list[num+1]:
            pik_stp_no = re.findall(r'^[0-9]',html_file_list[num-1]).pop()

    # Organize data into orders for regular BOL
    for num in range_of_list:
        if 'MABD' in html_file_list[num]: # num = 34

            # While our counter is less than total_po, check if our pick stop number
            # is the same as the pick stop numbers in the data, and then
            # categorize data accordingly into a huge list of dictionaries.
            # Also add load number and date into each dictionary.
            # Add 1 to counter, and 15 to num (the index) to move on to the next "set".

            count = 0
            while count < total_po:
                if pik_stp_no == html_file_list[num+1]:
                    order_list.append({
                        'po':html_file_list[num+4],
                        'dc':html_file_list[num+5],
                        'pkg':html_file_list[num+11],
                        'wt':html_file_list[num+12],
                        'mabd':html_file_list[num+15][:-9],
                        'ln':load_no,
                        'date':date.today().strftime('%m-%d-%Y')
                        })
                num += 15
                count += 1

    return order_list

# Run initialGui and assign two variables to it's output
html_file_path, dest_path = initialGui()

# Run readHtmlFile using those variables
html_file_list = readHtmlFile(html_file_path, dest_path)

# Run sortHtmlFile
order_list = sortHtmlFile(html_file_list)

print(order_list)