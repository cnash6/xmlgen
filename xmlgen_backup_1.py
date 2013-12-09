# Christian Nash/ Christian Tuchez
# 7/30/12
# XML Generator for testing purposes

from tkinter import *
import tkinter.messagebox as tkmb
import tkinter.filedialog as tkfd
import csv
import xml.etree.ElementTree as et
from time import gmtime, strftime
import os
from tkinter import ttk
import time


class GUI:
    

    def __init__(self):   # Init calls first window creator
        self.startPage()

    def startPage(self):
        self.rootwin = Tk()             
        self.rootwin.title("XML Generator")
        root = self.rootwin
        f = Frame(root, bg="#DBDBDB")
        f.pack(expand = 1, fill= BOTH)

        spacer = Label(f, text="", bg = "#DBDBDB", width=60)
        spacer.grid(row=1, columnspan=2)

        self.title = Label(f, text="XML Generator", font=("Helvetica", 16), bg="#DBDBDB")
        self.title.grid(row=2, column=0, columnspan=2)

        spacer = Label(f, text="", bg = "#DBDBDB")
        spacer.grid(row=3, columnspan=2)

        XMLLabel = Label(f, text="Select XML File Template", bg="#DBDBDB")
        XMLLabel.grid(row=4, column=0, columnspan=2)


        browse = Button(f, text="Browse", command=self.openXMLFile)
        browse.grid(row=5, column=0, pady=5, padx=5, sticky="e")

        self.XMLEntry = Entry(f, width=50)
        self.XMLEntry.insert(0,'. . .')
        self.XMLEntry.config(state='readonly')
        self.XMLEntry.grid(row=5, column=1, pady=5, sticky="w")

        CSVLabel = Label(f, text="Select a CSV File", bg="#DBDBDB")
        CSVLabel.grid(row=6, column=0, columnspan=2)

        browse = Button(f, text="Browse", command=self.openCSVFile)
        browse.grid(row=7, column=0, pady=5, padx=5, sticky='e')

        self.CSVEntry = Entry(f, width=50)
        self.CSVEntry.insert(0,'. . .')
        self.CSVEntry.config(state='readonly')
        self.CSVEntry.grid(row=7, column=1, pady=5, sticky="w")

        DLabel = Label(f, text="Select File Destination", bg="#DBDBDB")
        DLabel.grid(row=8, column=0, columnspan=2)

        browse = Button(f, text="Browse", command=self.getDirectory)
        browse.grid(row=9, column=0, pady=5, padx=5, sticky='e')

        self.DEntry = Entry(f, width=50)
        self.DEntry.insert(0,'. . .')
        self.DEntry.config(state='readonly')
        self.DEntry.grid(row=9, column=1, pady=5, sticky="w")

        cont = Button(f, text="Continue", command=self.doIt)
        cont.grid(row=10, pady=5, columnspan=2)

        self.isWorking = False
        test = Button(f, text="Test", command=self.progress)
        test.grid(row=11, pady=5, columnspan=2)

        menubar = Menu(root)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Exit", command=self.exitwin)
        menubar.add_cascade(label="File", menu=filemenu)
        helpmenu = Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About...", command=self.showHelp)
        menubar.add_cascade(label="Help", menu=helpmenu)
        root.config(menu=menubar)
	
        self.rootwin.mainloop()

    def showHelp(self):
        tkmb.showinfo("Help",
                      "XML Generator 1.0\nChristian Nash \n\n"
                      + "Instructions for use:\n\n"
                      + "Upload the desired XML file for replication. Then upload a CSV file with the following format:\n\n"
                      + "Parent/Child/DesiredField, Parent2/Child2/DesiredField2,\ndata1,\t\tdata2,\ndata1,\t\tdata2,\netc..."
                      + "\n\nThese CSV files are easily created by saving a excell spreadsheet in .csv format.\nSelect file destination and click \"Continue\" to auto-generate XML Files"
                      + "\n\n-Not intended for use outside of Manhattan Associates Inc."
                      )

    def exitwin(self):
        self.rootwin.destroy()
        

    def openXMLFile(self): #this method simply opens a file dialog box when the select file button is clicked. it allows teh user to select a file with which he/she wishes to use. it puts teh file path in the file entry field and sets it to readonly so teh user can copy it but not manipulate it
        self.XMLfname = tkfd.askopenfilename()
        if self.XMLfname:
            self.XMLEntry.config(state=NORMAL)
            self.XMLEntry.delete(0,END)
            self.XMLEntry.insert(0,self.XMLfname)
            self.XMLEntry.config(state='readonly')

    
    def openCSVFile(self):
        self.CSVfname = tkfd.askopenfilename()
        if self.CSVfname:
            self.CSVEntry.config(state=NORMAL)
            self.CSVEntry.delete(0,END)
            self.CSVEntry.insert(0,self.CSVfname)
            self.CSVEntry.config(state='readonly')

    def getDirectory(self):
        self.fDirectory = tkfd.askdirectory()
        if self.fDirectory:
            self.DEntry.config(state=NORMAL)
            self.DEntry.delete(0,END)
            self.DEntry.insert(0,self.fDirectory)
            self.DEntry.config(state='readonly')

    def doIt(self):
        try:
            if self.XMLfname[len(self.XMLfname)-4:] != '.xml':
                raise ValueError
            XMLfile = open(self.XMLfname,'r')
            myTree = et.parse(XMLfile)
            XMLfile.close()
            self.myRoot = myTree.getroot()
            self.filename = self.XMLfname[self.XMLfname.rfind('/')+1:len(self.XMLfname)-4]
        except:
            tkmb.showwarning("Error", "Error opening XML file.  Please verify file path and extension (.xml only)")
            return

        try:
            if self.CSVfname[len(self.CSVfname)-4:] != '.csv':
                raise ValueError
            CSVFile = open(self.CSVfname, 'r')
            self.csvdata = CSVFile.readlines()
            CSVFile.close()
            for x in range(len(self.csvdata)):
                self.csvdata[x] = self.csvdata[x].split(',')
                self.csvdata[x][len(self.csvdata[x])-1] = self.csvdata[x][len(self.csvdata[x])-1][:len(self.csvdata[x][len(self.csvdata[x])-1])-1]
        except:
            tkmb.showwarning("Error", "Error opening CSV file.  Please verify file path and extension (.csv only)")
            return

        if (tkmb.askyesno("Continue?", "Auto-Generate " + str(len(self.csvdata)-1) + " XML Files from template at \"" + self.fDirectory + "\"?") == False):
            return
        
        fields = self.csvdata[0]
        for x in fields:
            if self.myRoot.find(x) == None:
                tkmb.showwarning("Error", "CSV fields properly qualified.  Please verify that CSV column headers match XML elements")
                return

        dirname = os.path.join(self.fDirectory,strftime("AutoGen%m-%d-%Y %H.%M.%S",gmtime()))

        if not os.path.isdir(dirname):
            os.mkdir(dirname)

        self.rootwin2 = Tk()
        self.rootwin2.title("Progress")
        mw2 = self.rootwin2

        self.progress = ttk.Progressbar(mw2, orient="horizontal", length=400, mode="determinate")
        self.progress.pack()

        time.sleep(1)
    
        anum = 0
        for x in range(len(self.csvdata)-1):
            print(x)
            for y in range(len(self.csvdata[x+1])):
                print(y)
                self.myRoot.find(self.csvdata[0][y]).text = self.csvdata[x+1][y]
            anum+=1
            myTree.write(os.path.join(dirname, self.filename + "_" + str(anum) + ".xml"))

        tkmb.showinfo("Success", "Files successfully generated at " + dirname)

        time.sleep(1)

        self.rootwin2.destroy()

    def progress(self):
        if self.isWorking == False:
            self.isWorking = True
            self.rootwin2 = Tk()
            self.rootwin2.title("Progress")
            mw2 = self.rootwin2

            self.progress = ttk.Progressbar(mw2, orient="horizontal", 
                                        length=400, mode="determinate")
            self.progress.pack()

            cont = Button(mw2, text="step", command=self.step)
            cont.pack()

             
            
    def step(self):
        self.progress.step()
    
    def closeProgress(self):
        self.rootwin2.destroy()
        self.isWorking = False



        

GUI()
