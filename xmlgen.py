# @author Christian Nash
# @version 2.0 - 10/14/2013
# XML Generator for testing purposes

from tkinter import *
import tkinter.messagebox as tkmb
import tkinter.filedialog as tkfd
import csv
import xml.etree.ElementTree as et
from time import gmtime, strftime
import datetime
import os
from tkinter import ttk
import time
import shutil
import threading
from tkinter import tix
import subprocess
import math
import queue


class GUI:
    

    def __init__(self):   # Init calls first window creator
        self.startPage()

    def startPage(self):    # Sets up the main window
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


        self.browse1 = Button(f, text="Browse", command=self.openXMLFile)
        self.browse1.grid(row=5, column=0, pady=5, padx=5, sticky="e")
        
        self.XMLfname = ". . ."
        self.XMLEntry = Entry(f, width=50)
        self.XMLEntry.insert(0,self.XMLfname)
        #self.XMLEntry.config(state='readonly')
        self.XMLEntry.grid(row=5, column=1, pady=5, sticky="w")

        CSVLabel = Label(f, text="Select a CSV File", bg="#DBDBDB")
        CSVLabel.grid(row=6, column=0, columnspan=2)

        self.browse2 = Button(f, text="Browse", command=self.openCSVFile)
        self.browse2.grid(row=7, column=0, pady=5, padx=5, sticky='e')

        self.CSVfname = ". . ."
        self.CSVEntry = Entry(f, width=50)
        self.CSVEntry.insert(0,self.CSVfname)
        #self.CSVEntry.config(state='readonly')
        self.CSVEntry.grid(row=7, column=1, pady=5, sticky="w")

        DLabel = Label(f, text="Select File Destination", bg="#DBDBDB")
        DLabel.grid(row=8, column=0, columnspan=2)

        self.browse3 = Button(f, text="Browse", command=self.getDirectory)
        self.browse3.grid(row=9, column=0, pady=5, padx=5, sticky='e')

        self.fDirectory = ". . ."
        self.DEntry = Entry(f, width=50)
        self.DEntry.insert(0,self.fDirectory)
        #self.DEntry.config(state='readonly')
        self.DEntry.grid(row=9, column=1, pady=5, sticky="w")

        self.cont = Button(f, text="Continue", command=self.doIt)
        self.cont.grid(row=10, pady=5, columnspan=2)

        self.isWorking = True
        #test = Button(f, text="Test", command=self.test) ## Button used to quickly fill fields using sample files
        #test.grid(row=11, pady=5, columnspan=2)

        menubar = Menu(root)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Load Sample Set", command=self.test)
        filemenu.add_command(label="Exit", command=self.exitwin)
        menubar.add_cascade(label="File", menu=filemenu)
        helpmenu = Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About...", command=self.showHelp)
        menubar.add_cascade(label="Help", menu=helpmenu)
        root.config(menu=menubar)

        self.datalen=0
	
        self.rootwin.mainloop()

    def showHelp(self): ## Shows help box
        tkmb.showinfo("Help",
                      "XML Generator 1.1\nChristian Nash\ncnash3712@gmail.com \n\n"
                      + "Instructions for use:\n\n"
                      + "Upload the desired XML file for replication. Then upload a CSV file with the following format:\n\n"
                      + "Parent/Child/DesiredField, Parent2/Child2/DesiredField2,\ndata1,\t\tdata2,\ndata1,\t\tdata2,\netc..."
                      + "\n\nThese CSV files are easily created by saving a excell spreadsheet in .csv format.\nSelect file destination and click \"Continue\" to auto-generate XML Files"
                      + "\n\n*Not intended for use outside of Manhattan Associates Inc."
                      )

    def exitwin(self): # Exit the root window. Obviously
        self.rootwin.destroy()
        

    def openXMLFile(self): # Asks user for XML file and enters it into text box
        self.XMLfname = tkfd.askopenfilename() 
        
        if self.XMLfname:
            self.XMLEntry.delete(0,END)
            self.XMLEntry.insert(0,self.XMLfname)
    
    def openCSVFile(self): # Asks user for CSV file and enters it into text box
        self.CSVfname = tkfd.askopenfilename()
        
        if self.CSVfname:
            self.CSVEntry.delete(0,END)
            self.CSVEntry.insert(0,self.CSVfname)

    def getDirectory(self): # Asks user for a directory file and enters it into text box
        self.fDirectory = tkfd.askdirectory()
        if self.fDirectory:
            self.DEntry.delete(0,END)
            self.DEntry.insert(0,self.fDirectory)
 
    def doIt(self): # Begins the main function.  This preps the xml file generation, sets up folders, opens files, etc...
        self.XMLfname = self.XMLEntry.get()
        self.CSVfname = self.CSVEntry.get()
        self.fDirectory = self.DEntry.get()
        if self.XMLfname == ". . .":
            tkmb.showwarning("Error", "Error opening XML file.  Please verify file path and extension (.xml only)")
            return

        if self.CSVfname == ". . .":
            tkmb.showwarning("Error", "Error opening CSV file.  Please verify file path and extension (.csv only)")
            return

        if self.fDirectory == ". . .":
            tkmb.showwarning("Error", "Error accessing File Destination.  Please verify file destination")
            return

        try:
            if len(self.XMLfname) < 5 or self.XMLfname[len(self.XMLfname)-4:] != '.xml':
                raise ValueError
            XMLfile = open(self.XMLfname,'r')
            self.myTree = et.parse(XMLfile)
            XMLfile.close()
            self.myRoot = self.myTree.getroot()
            self.filename = self.XMLfname[self.XMLfname.rfind('/')+1:len(self.XMLfname)-4]
        except:
            tkmb.showwarning("Error", "Error opening XML file.  Please verify file path and extension (.xml only)")
            return
        
        try:
            if len(self.CSVfname) < 5 or self.CSVfname[len(self.CSVfname)-4:] != '.csv':
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
        
        if not os.path.isdir(self.DEntry.get()):
            tkmb.showwarning("Error", "File path does not exist!")
            return
        
        if (tkmb.askyesno("Continue?", "Auto-Generate " + str(len(self.csvdata)-1) + " XML Files from template at \"" + self.fDirectory + "\"?") == False):
            return

        self.browse1.config(state=DISABLED)
        self.browse2.config(state=DISABLED)
        self.browse3.config(state=DISABLED)
        self.cont.config(state=DISABLED) 

        self.fields = self.csvdata[0]
        for x in self.fields:
            if self.myRoot.find(x) == None:
                tkmb.showwarning("Error", "CSV fields not properly qualified.  Please verify that CSV column headers match XML elements")
                self.closeProg()
                return

        self.dirname = os.path.join(self.fDirectory,strftime("AutoGen%m-%d-%Y %H.%M.%S",gmtime()))

        if not os.path.isdir(self.dirname):
            os.mkdir(self.dirname)

        self.isWorking = True
        self.datalen = len(self.csvdata)
        self.anum = 1

        self.progressWin()
        
        self.timer()

        self.generate()

        #self.progress.config(value="100")
        
        #time.sleep(1)

        self.cancel.config(state=DISABLED)

        if self.isWorking:
            tkmb.showinfo("Success", "Files successfully generated at " + self.dirname)
        else:
            tkmb.showinfo("Failed", "Something done messed up.  Hopefully this just message just means you cancelled it")
            self.closeProg()
            return

        self.viewButton.config(state=NORMAL)
        self.closeProgress.config(state=NORMAL)

        self.rootwin2.lift()

        self.progress.config(value=str(len(self.csvdata)))
        
        
        
        
        
        

    def generate(self):
        
        for x in range(len(self.csvdata)-1):
            if not self.isWorking:
                shutil.rmtree(self.dirname)
                return
            for y in range(len(self.csvdata[x+1])):
                self.myRoot.find(self.csvdata[0][y]).text = self.csvdata[x+1][y]
            self.myTree.write(os.path.join(self.dirname, self.filename + "_" + str(self.anum) + ".xml"))
            self.progressLabel.config(text=("Generated "+str(self.anum)+" out of "+str(self.datalen-1)+" files"))
            self.anum+=1
            if self.anum < self.datalen:
                self.progress.step()
            self.end_time = time.time()
            self.timeLabel.config(text=str(datetime.timedelta(seconds=math.ceil((self.end_time - self.start_time)))))
            self.rootwin2.update()
        
        


    def timer(self):
        self.start_time = time.time()

    def progressWin(self):
        self.rootwin2 = Toplevel()
        self.rootwin2.title("Progress")
        self.rootwin2.protocol("WM_DELETE_WINDOW", self.closeProg)
        mw2 = self.rootwin2
        f = Frame(mw2, bd=20)
        f.pack()

        label = Label(f, text="Generating Files...")
        label.grid(row=0, columnspan=2, pady=10)

        self.progress = ttk.Progressbar(f, orient="horizontal", length=400, mode="determinate", maximum=len(self.csvdata))
        self.progress.grid(row=1, columnspan=2)

        self.progressLabel = Label(f, text=("Generated 0 out of "+str(self.datalen-1)+" files"))
        self.progressLabel.grid(row=2, column=0)

        self.timeLabel = Label(f, text="00:00")
        self.timeLabel.grid(row=3, column=0, )

        self.cancel = Button(f, text="Cancel", command=self.breakLoop)
        self.cancel.grid(row=2, column=1, rowspan=2, pady=10)

        ttk.Separator(f, orient=HORIZONTAL).grid(row=4, columnspan=2, sticky="ew")

        self.viewButton = Button(f, text="View Files", state=DISABLED, command=self.viewFiles)
        self.viewButton.grid(row=5, column=0, pady=10)

        self.closeProgress = Button(f, text="Close", state=DISABLED, command=self.closeProg)
        self.closeProgress.grid(row=5, column=1, pady=10)
        

    def viewFiles(self):
        subprocess.Popen('explorer '+self.dirname)

    def closeProg(self):
        self.rootwin2.destroy()
        self.browse1.config(state=NORMAL)
        self.browse2.config(state=NORMAL)
        self.browse3.config(state=NORMAL)
        self.cont.config(state=NORMAL)

    def breakLoop(self):
        self.isWorking = False

    def test(self):
        self.XMLfname = os.getcwd()+"/SAMPLE_FILES/SAMPLE_XML.xml"
        self.XMLEntry.delete(0,END)
        self.XMLEntry.insert(0,self.XMLfname)


        self.CSVfname = os.getcwd()+"/SAMPLE_FILES/SAMPLE_CSV_VERY_LARGE.csv"
        self.CSVEntry.delete(0,END)
        self.CSVEntry.insert(0,self.CSVfname)

        self.fDirectory = os.getcwd()+"\SAMPLE_FILES"
        self.DEntry.delete(0,END)
        self.DEntry.insert(0,self.fDirectory)

        
##        if self.isWorking == False:
##            self.isWorking = True
##            self.rootwin2 = Tk()
##            self.rootwin2.title("Progress")
##            mw2 = self.rootwin2
##
##            self.progress = ttk.Progressbar(mw2, orient="horizontal", 
##                                        length=400, mode="determinate")
##            self.progress.pack()
##
##            cont = Button(mw2, text="step", command=self.step)
##            cont.pack()

    ####unused
    def test2(self):
        threading.Thread.__init__(self)

    ####unused   
    def step(self):
        self.progress.step()



        

GUI()
