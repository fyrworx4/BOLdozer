# Modules for GUI
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

# Modules for HTML File Reading
import codecs
import io
from bs4 import BeautifulSoup as bs

# Modules for HTML File Parsing
import re
import xlrd
from datetime import date

# Modules for Mail Merge
from mailmerge import MailMerge
import os
import sys
import subprocess


def initialGui():

    # Tkinter settings
    root = Tk()
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



def readHtmlFile(html_file_path):
    
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
    
    # The sortHtmlFile function takes the info from html_file_list and
    # parses it into 

    # Define some variables
    range_of_list = range(len(html_file_list))
    order_list = []
    weight_list = []
    pkgs_list = []
    master_order_list = []
    today = date.today().strftime('%m-%d-%Y')


    # Extract load number, which is always the 2nd item
    global load_no
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
                        'po_num':html_file_list[num+4],
                        'dc':html_file_list[num+5],
                        'pkgs':html_file_list[num+11],
                        'weight':html_file_list[num+12],
                        'mabd':html_file_list[num+15][:-9],
                        'load_no':load_no,
                        'date':today
                    })
                num += 15
                count += 1

    # Map DCs to locations, add locations to order_list
    total_po_ours = len(order_list)
    excel_file = ('./dc-mappings.xlsx')
    wb = xlrd.open_workbook(excel_file)
    dc_mappings = wb.sheet_by_index(0)

    for i in range(total_po_ours):
        for num in range(dc_mappings.nrows):
            if int(order_list[i]['dc']) == int(dc_mappings.cell_value(num,0)):
                order_list[i].update({
                    'address':dc_mappings.cell_value(num,1),
                    'city_state_zip':dc_mappings.cell_value(num,2)
                })
    
    # Calculate total weight and total packages
    for i in range(total_po_ours):
        weight_list.append(int(order_list[i]['weight']))
        pkgs_list.append(int(order_list[i]['pkgs']))

    tw = sum(weight_list)
    tp = sum(pkgs_list)

    # Calculate total page numbers
    if type(total_po_ours / 8) == int:
        total_pages = int(total_po_ours/8)
    else:
        total_pages = int(((total_po_ours-(total_po_ours%8))/8)+1)

    # Organize data into master order list for master BOL
    for i in range(total_po_ours):

        # Check if beginning of series
        if i % 8 == 0:

            # Add page number, total pages, load number, and date for each page
            master_order_list.append({
                'p_n':str(int(i/8)+1),
                'to_p':str(total_pages),
                'load_no':str(load_no),
                'date':today
            })

            # Define some variables
            dict_no = int((i-(i%8))/8)
            distance = total_po_ours - i
            current_page = master_order_list[dict_no]

            # Add data to page dictionaries
            if distance >= 8:
                count = 0
                while count < 8:
                    current_page.update({
                        f'd{count}':str(order_list[dict_no+count]['dc']),
                        f'p{count}':str(order_list[dict_no+count]['po_num']),
                        f'c{count}':str(order_list[dict_no+count]['pkgs']),
                        f'w{count}':str(order_list[dict_no+count]['weight']),
                        f'm{count}':str(order_list[dict_no+count]['mabd'])
                    })
                    count += 1
                current_page.update({'tw':'','tp':''})
            else:
                count = 0
                while count < distance:
                    current_page.update({
                        f'd{count}':str(order_list[dict_no+count]['dc']),
                        f'p{count}':str(order_list[dict_no+count]['po_num']),
                        f'c{count}':str(order_list[dict_no+count]['pkgs']),
                        f'w{count}':str(order_list[dict_no+count]['weight']),
                        f'm{count}':str(order_list[dict_no+count]['mabd']),
                        'tw':str(tw),
                        'tp':str(tp)
                    })
                    count += 1

    return order_list, master_order_list


def mailMerge(order_list, master_order_list, dest_path):

    #directory = os.getcwd()

    if os.path.isfile(f'{dest_path}/Walmart BOL {load_no}.docx'):
        os.remove(f'{dest_path}/Walmart BOL {load_no}.docx')
    
    if os.path.isfile(f'{dest_path}/Walmart Master BOL Load ID {load_no}.docx'):
        os.remove(f'{dest_path}/Walmart Master BOL Load ID {load_no}.docx')


    bol = MailMerge('bol_template.docx')
    bol.merge_pages(order_list)
    bol.write(f'{dest_path}/Walmart BOL {load_no}.docx')

    master_bol= MailMerge('master_bol_template copy.docx')
    master_bol.merge_pages(master_order_list)
    master_bol.write(f'{dest_path}/Walmart Master BOL Load ID {load_no}.docx')

    if sys.platform == 'darwin':
        def openFolder(path):
            subprocess.Popen(['open', '--', path])
    elif sys.platform == 'linux2':
        def openFolder(path):
            subprocess.Popen(['xdg-open', '--', path])
    elif sys.platform == 'win32':
        def openFolder(path):
            subprocess.Popen(['explorer', path])
    
    openFolder(dest_path)

    

# Run initialGui and assign two variables to it's output
html_file_path, dest_path = initialGui()

# Run readHtmlFile using those variables
html_file_list = readHtmlFile(html_file_path)

# Run sortHtmlFile
order_list, master_order_list = sortHtmlFile(html_file_list)

# Run mailMerge
mailMerge(order_list, master_order_list, dest_path)