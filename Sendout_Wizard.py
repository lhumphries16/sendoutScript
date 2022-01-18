# -*- coding: utf-8 -*-
"""
Created on Wed Dec 22 13:59:22 2021

@author: trehu


"""

import pdfrw
import pandas as pd
import tkinter as tk
import webbrowser
import os

pdf_template = os.path.join('assets', 'sendoutRequestTemplate.pdf')#"assets/sendoutRequestTemplate.pdf"

ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

successPopupText = "PDF named as the influencer's name has been created in this folder and spreadsheet row has been copied to clipboard!"

def getFormFields(form):
    """
    Parameters
    ----------
    form : PDF Document
        Template form that is to be filled.

    Returns
    -------
    fieldsDict : Dict
        All fields of 'form' that are fillable, with empty strings as values.
    """
    keys = []
    fieldsDict = {}
    
    template_pdf = pdfrw.PdfReader(pdf_template)
    
    for page in template_pdf.pages:
        annotations = page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    keys.append(key)
                    
    for key in keys:
        fieldsDict[key] = ''
        
    return fieldsDict
                    
def getUserInput(fields):
    """
    Parameters
    ----------
    fields : dict
        Fields of the PDF document that are fillable, with empty strings as values.

    Returns
    -------
    inputDict : dict
        'fields' but propagated with user input
    """
    inputDict = fields
    
    
    for field in fields: 
        if field == 'Shipping Address':
            print("\n(Please seperate each part of the address with a ','. Ex: 123 Street, City, State, ZIP)")
        
        inputDict[field] = input(field + ': ')
        
    return inputDict

def fill_pdf(input_pdf_path, data_dict):
    """
    Parameters
    ----------
    input_pdf_path : str
        path for template PDF file to be filled
    output_pdf_path : str
        path for output PDF file
    data_dict : dict
        'fields' but propagated with user input

    Returns
    -------
    None.
    """
        
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    output_pdf_path = "filledPDF.pdf"
    
    for page in template_pdf.pages:
        annotations = page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in data_dict.keys():
                        if type(data_dict[key]) == bool:
                            if data_dict[key] == True:
                                annotation.update(pdfrw.PdfDict(
                                    AS=pdfrw.PdfName('Yes')))
                        else:
                            annotation.update(
                                pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                            )
                            annotation.update(pdfrw.PdfDict(AP=''))
        #Name file as influencers name
    if data_dict['Name']:
            influencerName = data_dict['Name'].split()
            output_pdf_path = os.path.join('Filled_Requests', influencerName[0] + influencerName[1] + '.pdf')  #"Filled_Requests/" + nameDate[0] + nameDate[1] + '.pdf'
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)
    template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
    #webbrowser.open(os.path(output_pdf_path))

def exportToExcel(data_dict):
    """
    Fills excel row with data.

    Parameters
    ----------
    data_dict : str
        dictionary of filled form fields

    Returns
    -------
    None.
    """
    
    sheetData = []
    productList = ''
    
    influencerName = data_dict['Name'].split()
    #First Name
    sheetData.append(influencerName[0])
    #Last Name
    sheetData.append(influencerName[1])
    sheetData.append(data_dict['Email Address'])
    sheetData.append(data_dict['phone number'])
    #Street
    sheetData.append(data_dict['Shipping Address'].split(', ')[0])
    #City
    sheetData.append(data_dict['Shipping Address'].split(', ')[1])
    #State
    sheetData.append(data_dict['Shipping Address'].split(', ')[2])
    #Zip
    #sheetData.append(data_dict['Shipping Address'].split(', ')[3])
    #Combine products into string
    for field in data_dict:
        if field.find('Product Name') != -1:
            productList += data_dict[field] + ', '
    #Store product string
    sheetData.append(productList)
    sheetData.append(data_dict['handle'])
    sheetData.append(data_dict['Instagram'])
    sheetData.append('yes')
    
    df = pd.DataFrame(sheetData).T
    df.to_clipboard("test.xlsx", header=False, index=False)

def center_window(width=300, height=200):
    # get screen width and height
    screen_width = master.winfo_screenwidth()
    screen_height = master.winfo_screenheight()

    # calculate position x and y coordinates
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    master.geometry('%dx%d+%d+%d' % (width, height, x, y))
    
def submit(event):
    entryList = [] #grab entries from boxes
    entryDict = formFieldsDict #all fields to be entered
    count = 0

    #Populate entry list with entries in order
    for entry in entryBoxes:
        if entry.get():
            entryList.append(entry.get())
        else:
            entryList.append("")
    #Populate fields from entry list
    for field in entryDict:
        entryDict[field] = entryList[count]
        count += 1
        
    #Error Messages
    if entryDict['Name'] == "":
        tk.messagebox.showerror(title="Error", message= "Please enter inflencer name.")
    if entryDict['Shipping Address'] == "":
        tk.messagebox.showerror(title="Error", message= "Please enter inflencer shipping address.")
    if entryDict['Request made by  date request made'] == "":
        tk.messagebox.showerror(title="Error", message= "Please enter employee name.")
    fill_pdf(pdf_template, entryDict)
    exportToExcel(entryDict)
    
    #Display success message
    tk.messagebox.showinfo(title="Success", message= successPopupText)


#Define tkinter master window
master = tk.Tk()
master.title("Sendout Wizard")

#Set icon to PF logo
master.iconbitmap(os.path.join('assets','icon.ico'))#'assets/icon.ico'

#Create scrollable canvas
canvas = tk.Canvas(master)

#Create frame
frame = tk.Frame(canvas)

#Add scrollbar in y-axis
scroll_y = tk.Scrollbar(master, orient="vertical", command=canvas.yview, bg="Red")

#Hardcoding number of rows and columns because I couldn't figure out how
#to do progrmattically
rows = 30
columns = 2
keyCount = -1
formFieldsDict = getFormFields(pdf_template)
keys = list(formFieldsDict.keys())
entryBoxes = []

#Setup entries as text vars for tkinter
for key in keys:
    entryBoxes.append(tk.StringVar())
        
#Title block
tk.Label(frame, text="Influencer Sendout Wizard",
         font=("Arial", 20), bg="White", fg="Red", relief="groove").pack(padx=55,pady=35, ipady=10, ipadx=10)
#Main Column
for row in range(len(keys)):       
    #Label
    tk.Label(frame, text=keys[keyCount]).pack()
    #Help labels for formatting
    if keys[keyCount] == "Shipping Address":
        tk.Label(frame, text="(123 Street, City, State, ZIP)").pack()
    """if keys[keyCount] == "Name":
        tk.Label(frame, text="(First Last 00/00/00)").pack()"""
    #Entry
    tk.Entry(frame, textvariable=entryBoxes[keyCount], width=60).pack(padx=55)
    keyCount += 1
    
  
# create a Submit Button and place into the root window
submitBtn = tk.Button(frame, text="Submit", fg="White",
                        bg="Red", command=submit).pack(ipady=10, ipadx=30, pady=20)
#submitBtn.grid(row=31, column=2, )

tk.Label(frame, text="Pointe Forest LLC, © 2021").pack()

# put the frame in the canvas
canvas.create_window(0, 0, anchor='nw', window=frame)

# make sure everything is displayed before configuring the scrollregion
canvas.update_idletasks()

canvas.configure(scrollregion=canvas.bbox('all'), 
                 yscrollcommand=scroll_y.set)
                 
canvas.pack(fill='both', expand=True, side='left')
scroll_y.pack(fill='y', side='right')

#Bind enter key to submit
master.bind('<Return>', submit)

center_window(490,600)
master.mainloop()
