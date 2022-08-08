#Limestone  Checker
# Sourav Bhatti
#6/1/2022



from tkinter import filedialog
from tkinter.filedialog import askopenfile, askopenfilename
import pandas as pd
import datetime
import ctypes
from tkinter import CENTER, END, Button, Checkbutton, Frame, IntVar, Menu, StringVar, font, simpledialog, ttk, messagebox
from tkinter.ttk import Combobox, Treeview
import requests
import json
import tkinter as tk
import pyperclip




root = tk.Tk()
root.geometry('1020x800+50+50')
root.title("Data Retriever V3.0.1 by Sourav Bhatti")
root.resizable(0,0)
ctypes.windll.shcore.SetProcessDpiAwareness(1)
root.configure(bg="grey16")

style = ttk.Style()
style.theme_use('clam')

text = ("Times new roman", 18, 'bold')
comboBoxFont =('Courier New', '15', "bold")
comboOption = StringVar()


qrCode =  StringVar()
tk.Label(root, text= "" ,  width = 15,bg="grey16" ).grid( column= 0, row = 0)
tk.Label(root, text= "Please select the Data you want to check." ,font=text,  width = 35, bg="grey16",fg='white' ).grid( column= 2, row = 0, pady=10)


combolist = Combobox(root, textvariable=comboOption, font=text, width=35, height=15, state='readonly' )
combolist.grid(column=2, row=1, pady=10)
combolist['values']=["MacId Duplicate Check","MacId","Firmware Version","Sim Card Carrier","Sim Card Number","IMEI Gateway", "Region", "Team Viewer ID"]
combolist.current(0)
root.option_add("*TCombobox*Listbox*Font", comboBoxFont)   # Change the font of the list

tk.Label(root, text= "Please Enter the QR code below." ,font= text, bg="grey16", fg='white').grid( column= 2, row = 2,pady=10)
qrEntry = tk.Entry(root, textvariable= qrCode , font=comboBoxFont, width = 40, bd=5, relief='ridge')
qrEntry.grid( column= 2, row = 5,pady=10)
tk.Label(root, text= "" ,  width = 40,height=3, bg="grey16").grid( column= 2, row = 8)









#####################################
#           TreeView                #
#####################################

style.configure("mystyle.Treeview", highlightthickness=50, bd=30, font=('Times new roman', 13, 'bold')) # Modify the font of the body
style.configure("mystyle.Treeview.Heading", font=('Times new roman', 15,'bold')) # Modify the font of the headings

displaydata = Treeview(root,style = 'mystyle.Treeview', columns=['QR CODE', 'Data Type','Data'], show='headings',
                        height=20, selectmode='extended', takefocus="Any")
displaydata.grid(column=2, row=9)
displaydata.column("# 1", anchor=CENTER, width=300)
displaydata.heading("# 1", text="QR CODE")
displaydata.column("# 2", anchor=CENTER, width=200)
displaydata.heading("# 2", text="Data Type")
displaydata.column("# 3", anchor=CENTER, width=300)
displaydata.heading("# 3", text="Data")

listLast2Mac = []

def getMacId():

    url = "http://vmprdate.eastus.cloudapp.azure.com:9000/api/v1/manifest/?qrcode=" + qrCode.get()
    r = requests.get(url)
    rawjson = json.loads(r.text)
    
    
    if rawjson['error'] == False:
        data = rawjson['data']


        if comboOption.get() == "MacId Duplicate Check" :
            
            macid = data[0]['macid']      
            last2Mac = macid[-2:]
            if last2Mac in listLast2Mac:
                messagebox.showerror("Duplicate Mac Id", f"Do not ship Milestone with QR code {qrCode.get()}")
                status = "Duplicate"
            else:
                listLast2Mac.append(last2Mac)
                status = "Good to ship"

            print(last2Mac)
    
        if comboOption.get() == "MacId" :
            macid = data[0]['macid']      # Very Important Line

    
        elif comboOption.get() =="Firmware Version" :
            macid = data[0]['fw_version'] 

        elif comboOption.get() =="IMEI Gateway" :
            macid = data[0]['IMEInumber'] 

        elif comboOption.get() =="Sim Card Carrier" :
            macid = data[0]['comment'] 
            string = str(macid)
            listOfString = string.split("|")
            macid = listOfString[1]
        
        elif comboOption.get() =="Team Viewer ID" :
            macid = data[0]['comment'] 
            string = str(macid)
            listOfString = string.split("|")
            macid = listOfString[2]
        
        elif comboOption.get() =="Region" :
            macid = data[0]['comment'] 
            string = str(macid)
            listOfString = string.split("|")
            macid = listOfString[9]
        
        elif comboOption.get() =="Sim Card Number" :
            macid = data[0]['SIMcard'] 
            
        
                
        if macid != "None" and macid != "":
            # macids.append(macid)
            
            displaydata.insert('', END,
                    values=[qrCode.get(),comboOption.get(), macid])
        else:
            messagebox.showerror("Macid not found for ",qrCode.get())
        serverResponse = json.dumps(data)

        tstamp = datetime.datetime.today().strftime("%m/%d/%y")
        with open("Log.csv", "a", newline='') as file:
            data = {"Date":[f"{tstamp}"],"QR code":[f"{qrCode.get()}"],"Data Type":[f"{comboOption.get()}"], "Data Value":[f'{macid}'], "Status":[f'{status}'] }
            df = pd.DataFrame(data)
            df.to_csv(file, index=False, header=False )

    qrEntry.delete(0,END)

    
root.bind('<Return>', lambda event: getMacId())

def copytree():

    root.clipboard_clear()
    if displaydata.focus():
        curItem = displaydata.focus()
        root.clipboard_append(f" (QR CODE) {displaydata.item(curItem)['values'][0]} | (MACID) - {displaydata.item(curItem)['values'][1]}\n")
    else:
        for items in displaydata.get_children():
            #current_selection = displaydata.focus()
            root.clipboard_append(f" (QR CODE) {displaydata.item(items)['values'][0]} | (MACID) - {displaydata.item(items)['values'][1]}\n" )
        
        
    messagebox.showinfo("Please paste before closing program.", root.clipboard_get())


def readcsv():
    

    csvfile = filedialog.askopenfilename()
    df = pd.read_excel(csvfile)
    
    for value in df:

        url = "http://vmprdate.eastus.cloudapp.azure.com:9000/api/v1/manifest/?qrcode=" + value
        r = requests.get(url)
        rawjson = json.loads(r.text)

        

        if rawjson['error'] == False:
            data = rawjson['data']


def batchMacId():

    csvfile = filedialog.askopenfilename()
    df = pd.read_excel(csvfile)
        
    for index in range(len(df)):

        # url = "http://vmprdate.eastus.cloudapp.azure.com:9000/api/v1/manifest/?qrcode=" + value
        # r = request.get(url)
        # rawjson = json.loads(r.text)

        qrcode = df["QR Code"].iloc[index]
        url = "http://vmprdate.eastus.cloudapp.azure.com:9000/api/v1/manifest/?qrcode=" + df["QR Code"].iloc[index]
        r = requests.get(url)
        rawjson = json.loads(r.text)
        data = rawjson["data"]
        macid = data[0]["macid"]
        print(f"{qrcode} | {macid}")

  


#################################################################
    # copiedString = ''
    # for line in displaydata.get_children():
        
    #     for value in displaydata.item(line)['values']:
    #         copiedString += f"{value}, "

    # pyperclip.copy(copiedString)

    # messagebox.showinfo("Copied", pyperclip.paste())
#################################################################

tk.Label(root, text= "" ,  width = 15,bg="grey16" ).grid( column= 2, row = 10)
tk.Label(root, text= "" ,  width = 15,bg="grey16" ).grid( column= 2, row = 11)

Button(root, text='Copy',width=15, command= lambda :copytree()).grid(column=2, row= 12)

root.mainloop()


