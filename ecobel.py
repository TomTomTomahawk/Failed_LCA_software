# --clean -y -n "ecobel" --add-data="database\database.xlsx;database" ecobel.py


#--------------------------DATABASE TREATMENT

import xlrd
import numpy as np

"""

book = xlrd.open_workbook('database.xlsx')
print(book.nsheets)
print(book.sheet_names())
sheet=book.sheet_by_index(0)
#lines then columns, starts at 0
cell = sheet.cell(3,0)
print(cell.value)

categories=[]
options=[]

for i in range (rows):
    categorycell = sheet.cell(i,0)
    if categorycell.value != "":
        categories.append(categorycell.value)
    
    else:
        
    
        options=[]
        optioncell = sheet.cell(i,1)
        options.append(optioncell)
    
    if nextcell.value =="":
        options.append(nextcell.value)
        
            
print(categories)

ingredients=[]

rows=842

for i in range (rows):
    cell = sheet.cell(i,1)
    ingredients.append(str(cell.value))
    
carbon_impacts=[]

for i in range (rows):
    cell = sheet.cell(i,4)
    carbon_impacts.append(str(cell.value))

"""

#--------------------------INITIATE TKINTER

import tkinter as tk
from tkinter import ttk
from tkinter import *
import tkinter.scrolledtext as tkscrolled

from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure



root = Tk()
root.title("Ecobel")
#root.geometry('{}x{}'.format(1150, 600))

w, h = float(root.winfo_screenwidth()), float(root.winfo_screenheight())
#root.overrideredirect(True)
#root.geometry("%dx%d+0+0" % (w, h))


root.wm_state('zoomed')
#root.overrideredirect(True)
root.attributes('-topmost', True)

#m = root.maxsize()
#root.geometry('{}x{}+0+0'.format(*m))


# create all of the main containers
top_frame = Frame(root,bg='medium sea green',width=450, height=50, pady=3)
center = Frame(root,bg='black',width=50, height=400, pady=3)
btm_frame = Frame(root,bg='medium sea green',width=450, height=50, pady=3)

# layout all of the main containers
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

top_frame.grid(row=0, sticky="ew")
center.grid(row=1, sticky="nsew")
btm_frame.grid(row=3, sticky="ew")

# create the center widgets
center.grid_rowconfigure(0, weight=1)
center.grid_columnconfigure(1, weight=1)

ctr_left = Frame(center, bg='RoyalBlue1',width=2.8*w/10, height=190,pady=0)
ctr_left2 = Frame(center, bg='RoyalBlue1',width=2.8*w/10, height=190,pady=0)
#ctr_mid = Frame(center, bg='steelblue2',width=1.8*w/10, height=190,pady=0)
ctr_right = Frame(center, bg='tomato2', width=4.4*w/10, height=190,pady=0)

ctr_left.grid(row=0, column=0, sticky="ns")
ctr_left.grid_propagate(0)
ctr_left2.grid(row=0, column=1, sticky="ns")
ctr_left2.grid_propagate(0)
#ctr_mid.grid(row=0, column=2, sticky="nsew")
ctr_left2.grid_propagate(0)
ctr_right.grid(row=0, column=3, sticky="ns")
ctr_right.grid_propagate(0)

#seperate the left center frame in top and bottom

ctr_left_top = Frame(ctr_left, bg='purple',width=300, height=80)
ctr_left_btm = Frame(ctr_left, bg='pink',width=300, height=400)

ctr_left_top.grid(row=0, column=0, sticky="ns")
ctr_left_btm.grid(row=1, column=0, sticky="nsew")

#seperate the left center bottom frame in left and right

ctr_left_btm_left = Frame(ctr_left_btm, bg='yellow',width=150, height=400)
ctr_left_btm_mid = Frame(ctr_left_btm, bg='yellow',width=150, height=400)
ctr_left_btm_right = Frame(ctr_left_btm, bg='red',width=150, height=400)
ctr_left_btm_right2 = Frame(ctr_left_btm, bg='red',width=150, height=400)

ctr_left_btm_left.grid(row=0, column=0, sticky="ns")
ctr_left_btm_mid.grid(row=0, column=1, sticky="ns")
ctr_left_btm_right.grid(row=0, column=2, sticky="nsew")
ctr_left_btm_right2.grid(row=0, column=3, sticky="nsew")




#seperate the left center frame in top and bottom

ctr_left2_top = Frame(ctr_left2, bg='purple',width=300, height=80)
ctr_left2_btm = Frame(ctr_left2, bg='pink',width=300, height=400)

ctr_left2_top.grid(row=0, column=0, sticky="ns")
ctr_left2_btm.grid(row=1, column=0, sticky="nsew")

#seperate the left center bottom frame in left and right

ctr_left2_btm_left = Frame(ctr_left2_btm, bg='yellow',width=150, height=400)
ctr_left2_btm_mid = Frame(ctr_left2_btm, bg='yellow',width=150, height=400)
ctr_left2_btm_right = Frame(ctr_left2_btm, bg='red',width=150, height=400)
ctr_left2_btm_right2 = Frame(ctr_left2_btm, bg='red',width=150, height=400)

ctr_left2_btm_left.grid(row=0, column=0, sticky="ns")
ctr_left2_btm_mid.grid(row=0, column=1, sticky="ns")
ctr_left2_btm_right.grid(row=0, column=2, sticky="nsew")
ctr_left2_btm_right2.grid(row=0, column=3, sticky="nsew")

#--------------------------TOP FRAME

presentation_label = Label(top_frame, text='To calculate the environmental impact of your food, add ingredients. If it\'s a meal and you don\'t know its mass, type 0.6kg', background = 'medium sea green', foreground = "black")

presentation_label.grid(row=0, columnspan=3)





#--------------------------MIDDLE FRAME

#--------------------------LEFT MIDDLE FRAME 1

import sys

if getattr(sys, 'frozen', False):
    database = xlrd.open_workbook(os.path.join(sys._MEIPASS,'database/database.xlsx'))
else:
    database = xlrd.open_workbook('database/database.xlsx')


#database = xlrd.open_workbook('database.xlsx')
databasesheet=database.sheet_by_index(0)
#lines then columns, starts at 0

ingredients=[]

rows=databasesheet.nrows

for i in range (rows):
    cell = databasesheet.cell(i,1)
    ingredients.append(str(cell.value))
    
carbon_impacts=[]

for i in range (rows):
    cell = databasesheet.cell(i,4)
    carbon_impacts.append(str(cell.value))


units=[]
for i in range (rows):
    cell = databasesheet.cell(i,3)
    units.append(str(cell.value))


#lines then columns, starts at 0
dicts={}
categories=[]
index=[]

for i in range(databasesheet.nrows):
    cell = databasesheet.cell(i,0)
    if cell.value != "":
        categories.append(cell.value)
        index.append(i-1)

index=index+[databasesheet.nrows]



options=[]

for i in range(databasesheet.nrows):
    cell = databasesheet.cell(i,1)
    if cell.value != "":
        options.append(cell.value)


for i in range(len(index)-1):
    if i == 0:
        dicts[categories[i]]=options[:index[i+1]]
    else:        
        dicts[categories[i]]=options[index[i]:index[i+1]]



    
    

def addBox():

    def boxy1(event):
        
        CategorySelected=ent0.get()
                
        ent1['values']=(dicts[CategorySelected])
    
    def unitDisplayer1(event):
        unit1.delete(0.0,END)
        choice=str(ent1.get())
        for i in range(rows):
            if choice==ingredients[i]:
                unit=units[i]
        message=unit
        unit1.insert(0.0,message)


    ent0 = ttk.Combobox(ctr_left_btm_left, values=categories)
    ent0.pack(side=TOP)
    ent0.current(0)
    ent0.bind("<<ComboboxSelected>>", boxy1)
    
    
    
    ent1 = ttk.Combobox(ctr_left_btm_mid,values=[''])
    ent1.pack(side=TOP)
    ent1.current(0)
    ent1.bind("<<ComboboxSelected>>", unitDisplayer1)


    ent2 = Entry(ctr_left_btm_right,font="Arial 11",width=5)
    
    ent2.pack(side=TOP)


    unit1 = Text(ctr_left_btm_right2,font="Arial 11",width=5,height=1)
    
    unit1.pack(side=TOP)


    





    all_entries.append( (ent1, ent2) )


def showEntries():
    display.delete(0.0,END)
    for number, (ent1, ent2) in enumerate(all_entries):
        ingredient=str(ent1.get())
        for i in range(rows):
            if ingredient==ingredients[i]:
                impact=carbon_impacts[i]
        message="ingredient" + str(number) + ": " + str(ent1.get()) + str(ent2.get()) + impact + '\n'
        print (message)
        display.insert(0.0,message)

all_entries = []

Label(ctr_left_top, text='Product 1').grid(row=0,column=0)

addboxButton = Button(ctr_left_top, text='<Add constituent>', fg="Red", command=addBox)
addboxButton.grid(row=1, column=0)

Label(ctr_left_btm_left, text='Category').pack(side=TOP)
Label(ctr_left_btm_mid, text='Constituent').pack(side=TOP)

Label(ctr_left_btm_right, text='%').pack(side=TOP)
Label(ctr_left_btm_right2, text='Unit').pack(side=TOP)

Label(ctr_left_top,text='total mass in kg').grid(row=2, column=0)

foodmass = Entry(ctr_left_top,font="Arial 11",width=5)
foodmass.grid(row=2, column=1)



#--------------------------LEFT MIDDLE FRAME 2

Label(ctr_left2_top, text='Product 2').grid(row=0,column=0)

all_entries2 = []

def addBox2():

    def boxy2(event):
        
        CategorySelected=ent.get()
                
        ent3['values']=(dicts[CategorySelected])

    def unitDisplayer2(event):
        unit2.delete(0.0,END)
        choice=str(ent3.get())
        for i in range(rows):
            if choice==ingredients[i]:
                unit=units[i]
        message=unit
        unit2.insert(0.0,message)
    
    ent = ttk.Combobox(ctr_left2_btm_left, values=categories)
    ent.pack(side=TOP)
    ent.current(0)
    ent.bind("<<ComboboxSelected>>", boxy2)
    
    
    
    ent3 = ttk.Combobox(ctr_left2_btm_mid,values=[''])
    ent3.pack(side=TOP)
    ent3.current(0)
    ent3.bind("<<ComboboxSelected>>", unitDisplayer2)


    ent4 = Entry(ctr_left2_btm_right,font="Arial 11",width=5)
    
    ent4.pack(side=TOP)


    unit2 = Text(ctr_left2_btm_right2,font="Arial 11",width=5,height=1)
    
    unit2.pack(side=TOP)


    all_entries2.append( (ent3, ent4) )




def showEntries():
    display.delete(0.0,END)
    for number, (ent3, ent4) in enumerate(all_entries):
        ingredient=str(ent3.get())
        for i in range(rows):
            if ingredient==ingredients[i]:
                impact=carbon_impacts[i]
        message="ingredient" + str(number) + ": " + str(ent3.get()) + str(ent4.get()) + impact + '\n'
        print (message)
        display.insert(0.0,message)


addboxButton2 = Button(ctr_left2_top, text='<Add constituent>', fg="Red", command=addBox2)
addboxButton2.grid(row=1, column=0)

Label(ctr_left2_btm_left, text='Category').pack(side=TOP)
Label(ctr_left2_btm_mid, text='Constituent').pack(side=TOP)

Label(ctr_left2_btm_right, text='%').pack(side=TOP)
Label(ctr_left2_btm_right2, text='Unit').pack(side=TOP)

Label(ctr_left2_top,text='total mass in kg').grid(row=2, column=0)

foodmass2 = Entry(ctr_left2_top,font="Arial 11",width=5)
foodmass2.grid(row=2, column=1)


#--------------------------CENTER MIDDLE FRAME







#--------------------------RIGHT MIDDLE FRAME

showButton = Button(ctr_right, text='Show all impacts', command=lambda:[showEntries(),plot()])
showButton.grid(row=0,column=0)

quitButton = tk.Button(ctr_right, text="QUIT", fg="red",command=root.destroy).grid(row=1,column=0)

display=tkscrolled.ScrolledText(ctr_right,width=72,height=7,wrap=WORD)
display.grid(row=3,column=0)

def plot():
    fig = Figure(figsize=(4,5), dpi=96)
    ax = fig.add_subplot(111)
    
    if foodmass.get() != "":
        impacts=[]
        labels=[]
        percentage=[]
        for number, (ent1, ent2) in enumerate(all_entries):
            ingredient=str(ent1.get())
            if str(ent1.get()) != "":
                labels.append(str(ent1.get()))
                percentage.append(float(ent2.get())/100)
            for i in range(rows):
                if ingredient==ingredients[i] and ingredient != "":
                    impact=carbon_impacts[i]
                    impacts.append(float(impact))
        objectmass=float(foodmass.get())    
        display.insert(0.0,str(len(percentage))+'  '+ str(len(impacts))+'  '+ str(objectmass))

        finalimpacts=[]
    
        for i in range(len(impacts)):
            finalimpacts.append(impacts[i]*percentage[i]*objectmass)
        


        for i in range(len(finalimpacts)):
            ax.bar(0,(finalimpacts[i]),bottom=np.sum(finalimpacts[:i]))
        
        labels12=labels       
        ax.set_xticks([0])
        ax.set_xticklabels(['Product 1'])
        ax.legend(labels12)
        
        
    if foodmass2.get() != "":


        impacts2=[]
        labels2=[]
        percentage2=[]
        for number, (ent3, ent4) in enumerate(all_entries2):
            ingredient2=str(ent3.get())
            if str(ent3.get()) != "":
                labels2.append(str(ent3.get()))
                percentage2.append(float(ent4.get())/100)
            for i in range(rows):
                if ingredient2==ingredients[i] and ingredient2 != "":
                    impact2=carbon_impacts[i]
                    impacts2.append(float(impact2))
        objectmass2=float(foodmass2.get())    
        display.insert(0.0,str(len(percentage2))+'  '+ str(len(impacts2))+'  '+ str(objectmass2))

        finalimpacts2=[]

    
        for i in range(len(impacts2)):
            finalimpacts2.append(impacts2[i]*percentage2[i]*objectmass2)
   
        for i in range(len(finalimpacts2)):
            ax.bar(1,(finalimpacts2[i]),bottom=np.sum(finalimpacts2[:i]))
        
        if foodmass.get() != "":
            labels12=labels+labels2
            ax.set_xticklabels(['Product 1','Product 2'])
        else:
            labels12=labels2
            ax.set_xticklabels(['Product 2'])
        ax.set_xticks([0,1])
        ax.legend(labels12)
    
    
    ax.set_ylabel('kg CO2 eq')

    fig.tight_layout()
    
    graph = FigureCanvasTkAgg(fig,ctr_right)
    canvas = graph.get_tk_widget()
    canvas.grid(row=2,column=0)


#--------------------------BOTTOM FRAME


credit_label = Label(btm_frame, text='Thomas Benoît Le Varlet - 2019',background = 'medium sea green', foreground = "black")

credit_label.pack(side=RIGHT)

root.mainloop()