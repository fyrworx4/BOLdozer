# Modules for GUI
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

# Modules for HTML File Parsing
import codecs
import io
from bs4 import BeautifulSoup as bs

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
        html_file_label = Label(root, text=f'Chosen file: {html_file_path}').grid(row=2, column=0)
        #print(html_file)

    def getDest():
        global dest_path
        dest_path = filedialog.askdirectory(initialdir='/', title='Open Directory')
        dest_path_label = Label(root, text=f'Chosen directory: {dest_path}').grid(row=5, column=0)
        #print(dest_path)

    # Very basic user input validation
    def checkInputValidation():
        try:
            html_file_path
            dest_path
            root.quit()
        except NameError:
            retry = messagebox.showwarning(title='Error', message="You didn\'t select anything!")

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

# Run initialGui and assign two variables to it's output
html_file_path, dest_path = initialGui()
# Run readHtmlFile using those variables
readHtmlFile(html_file_path, dest_path)