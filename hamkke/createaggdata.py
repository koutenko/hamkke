#!/usr/bin/env python
"""
Module Name
===========
 $ createaggdata.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Shell Syntax
============
 N/A

Description
===========
This module allows for the creation of aggregate data files.
"""

import sys
import os
import Tkinter as tk
import hamkke as ra
import time
import sqlite3
import re
from sqlite3 import Error
from urlparse import urljoin
from urllib import pathname2url
from subprocess import check_call, CalledProcessError, STDOUT
from pyPdf import PdfFileReader, PdfFileWriter
from wkhtmltopdf import WKhtmlToPdf
import hamkke as ra
from hamkke.cahps import constants as const
from hamkke.mcahps import constants as mconst

class AggregateData(tk.Frame):
    def __init__(self, frame):
        '''
        This module allows python to update the database and then run a simple
        test to confirm accurate entry.
        '''
        # Initialise the base Tkinter frame
        tk.Frame.__init__(self)

        # Pull data from the 201X CAHPS schedule
        # Data = ra.get_sqlite_data(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\Schedule\CAHPS16.db",
        #                         'Projects')

        # Create a parent frame to display the module
        top_frame = frame

        # Process to clear anything hanging around before beginning anew
        for widget in top_frame.winfo_children():
            widget.destroy()
        
        # Variable definitions for the Entry boxes
        self.aggvar = tk.StringVar()

        self.idvar_1 = tk.StringVar()
        self.idvar_2 = tk.StringVar()
        self.idvar_3 = tk.StringVar()
        self.idvar_4 = tk.StringVar()
        self.idvar_5 = tk.StringVar()
        self.idvar_6 = tk.StringVar()
        self.idvar_7 = tk.StringVar()
        self.idvar_8 = tk.StringVar()
        self.idvar_9 = tk.StringVar()
        self.idvar_10 = tk.StringVar()
        self.idvar_11 = tk.StringVar()
        self.idvar_12 = tk.StringVar()
        self.idvar_13 = tk.StringVar()
        self.idvar_14 = tk.StringVar()
        self.idvar_15 = tk.StringVar()
        self.idvar_16 = tk.StringVar()
        self.idvar_17 = tk.StringVar()
        self.idvar_18 = tk.StringVar()
        self.idvar_19 = tk.StringVar()
        self.idvar_20 = tk.StringVar()

        self.surveyvar = tk.StringVar()
        self.planvar = tk.StringVar()
        self.q1var = tk.StringVar()
        self.productvar = tk.StringVar()
        self.error_msg = tk.StringVar()
        self.outputvar = tk.StringVar()

        # Validation tracer
        #self.idvar.trace('w', self.validate_input)

        # Create child frame within which everything is displayed
        self.main_frame = tk.Frame(frame, width=800, height=600)
        self.main_frame.config(bg="#424242")
        self.main_frame.grid(row=0, column=0)
        self.main_frame.grid_propagate(0)
        self.main_frame.update()
        self.main_frame.isgridded = False # dynamically add "isgridded" attribute

        # Error label appears as needed
        self.error_label = tk.Label(self.main_frame,
                                    textvariable=self.error_msg,
                                    font="Calibri 11",
                                    fg="#E8483B",
                                    bg="#424242")
        
        # Bring in the text output box
        self.text_ui()      


        # ------------- H E A D E R -------------
        # MODULE TITLE
        mod_title_label = tk.Label(self.main_frame,
                            text="- C R E A T E  A G G R E G A T E  D A T A -",
                            font="Calibri 18 bold",
                            fg="#E8B93B",
                            bg="#424242")
        mod_title_label.place(x=400, y=12, anchor="center")
        # MODULE INSTRUCTIONS
        instructions_label = tk.Label(self.main_frame, 
                               text="This module allows python to create aggregate data files by copying " + \
                                    "\nthe text data files of the projects to be merged " + \
                                     "\ninto a new directory before combining them.",
                               font="Calibri 11",
                               fg="#CCCCCC",
                               bg="#424242")
        instructions_label.place(x=400, y=65, anchor="center")

        # ----------- S E L E C T  B U T T O N S -----------
        modes = [
            ("New Entry", "1"),
            ("Update Entry", "2"),

        ]

        global vari
        vari = tk.IntVar()

        # CAHPS
        self.r1 = tk.Radiobutton(self.main_frame, 
                            text="CAHPS", 
                            variable=vari, 
                            value=1, 
                            indicatoron=0,
                            relief="groove",
                            font="Calibri 11",
                            fg="#CCCCCC",
                            bg="#424242",
                            width=10,
                            command=lambda: self.select())
        # MCAHPS
        self.r2 = tk.Radiobutton(self.main_frame,
                            text="MCAHPS",
                            variable=vari,
                            value=2,
                            indicatoron=0,
                            relief="groove",
                            font="Calibri 11",
                            fg="#CCCCCC",
                            bg="#424242",
                            width=10,
                            command=lambda: self.select())

        self.r1.place(x=47, y=160)
        self.r2.place(x=47, y=190)

        # -------------- I N F O  F I E L D S --------------
        self.id_label = tk.Label(self.main_frame, text="Project IDs: ", font="Calibri 11 bold", fg="#CCCCCC", bg="#424242")
        self.id_label.place(x=400, y=100)
        
        # FILE 01 PROJECT ID
        self.id_field_1 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_1)
        self.id_field_1.place(x=280, y=130)
        self.id_field_1.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 02 PROJECT ID
        self.id_field_2 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_2)
        self.id_field_2.place(x=280, y=160)
        self.id_field_2.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 03 PROJECT ID
        self.id_field_3 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_3)
        self.id_field_3.place(x=280, y=190)
        self.id_field_3.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 04 PROJECT ID
        self.id_field_4 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_4)
        self.id_field_4.place(x=280, y=220)
        self.id_field_4.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 05 PROJECT ID     
        self.id_field_5 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_5)
        self.id_field_5.place(x=280, y=250)
        self.id_field_5.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 06 PROJECT ID       
        self.id_field_6 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_6)
        self.id_field_6.place(x=360, y=130)
        self.id_field_6.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 07 PROJECT ID       
        self.id_field_7 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_7)
        self.id_field_7.place(x=360, y=160)
        self.id_field_7.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 08 PROJECT ID       
        self.id_field_8 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_8)
        self.id_field_8.place(x=360, y=190)
        self.id_field_8.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 09 PROJECT ID       
        self.id_field_9 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_9)
        self.id_field_9.place(x=360, y=220)
        self.id_field_9.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 10 PROJECT ID       
        self.id_field_10 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_10)
        self.id_field_10.place(x=360, y=250)
        self.id_field_10.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 11 PROJECT ID
        self.id_field_11 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_11)
        self.id_field_11.place(x=440, y=130)
        self.id_field_11.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 12 PROJECT ID
        self.id_field_12 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_12)
        self.id_field_12.place(x=440, y=160)
        self.id_field_12.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 13 PROJECT ID
        self.id_field_13 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_13)
        self.id_field_13.place(x=440, y=190)
        self.id_field_13.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 14 PROJECT ID
        self.id_field_14 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_14)
        self.id_field_14.place(x=440, y=220)
        self.id_field_14.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 05 PROJECT ID     
        self.id_field_15 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_15)
        self.id_field_15.place(x=440, y=250)
        self.id_field_15.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 06 PROJECT ID       
        self.id_field_16 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_16)
        self.id_field_16.place(x=520, y=130)
        self.id_field_16.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 07 PROJECT ID       
        self.id_field_17 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_17)
        self.id_field_17.place(x=520, y=160)
        self.id_field_17.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 08 PROJECT ID       
        self.id_field_18 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_18)
        self.id_field_18.place(x=520, y=190)
        self.id_field_18.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 09 PROJECT ID       
        self.id_field_19 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_19)
        self.id_field_19.place(x=520, y=220)
        self.id_field_19.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # FILE 10 PROJECT ID       
        self.id_field_20 = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar_20)
        self.id_field_20.place(x=520, y=250)
        self.id_field_20.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

         # PRODUCT LINE
        aggregate_label = tk.Label(self.main_frame, text="Aggregate File\nProject ID: ", font="Calibri 11", fg="#CCCCCC", bg="#424242")
        aggregate_label.place(x=160, y=145)
        
        self.aggregate_field = tk.Entry(self.main_frame, relief="flat", textvariable=self.aggvar)
        self.aggregate_field.place(x=170, y=190)
        self.aggregate_field.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # -------------- B U T T O N S --------------
        # GENERATE DIRECTORY BUTTON
        self.create_agg_button = tk.Button(self.main_frame,
                                  text="Create\nDirectory",
                                  command=lambda: self.create_agg_directory(),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=10)
        self.create_agg_button.place(x=163, y=217)
        self.create_agg_button.config(highlightbackground="dark gray", highlightthickness=1, state="disabled")

        # CREATE BUTTON
        self.create_button = tk.Button(self.main_frame,
                                  text="Create",
                                  command=lambda: self.create_aggregate(files),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=7)
        self.create_button.place(x=670, y=152)
        self.create_button.config(highlightbackground="dark gray", highlightthickness=1, state="disabled")

        # CLEAR BUTTON
        self.clear_button = tk.Button(self.main_frame,
                                  text="Clear",
                                  command=lambda: self.clear_text(),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=7)
        self.clear_button.place(x=670, y=187)
        self.clear_button.config(highlightbackground="dark gray", highlightthickness=1)

        # CLEAR ALL BUTTON
        self.clear_all_button = tk.Button(self.main_frame,
                                  text="Clear All",
                                  command=lambda: self.clear_all(),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=7)
        self.clear_all_button.place(x=670, y=222)
        self.clear_all_button.config(highlightbackground="dark gray", highlightthickness=1)

        # CONFIRM FILES BUTTON
        self.confirm_button = tk.Button(self.main_frame,
                                  text="Confirm\nFiles",
                                  command=lambda: self.confirm_files(),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=7,
                                  height=7)
        self.confirm_button.place(x=600, y=132)
        self.confirm_button.config(highlightbackground="dark gray", highlightthickness=1, state="disabled")

        
    # -------------- C O P Y  P R O C E S S --------------
    def select(self):
        global choice
        global basedir
        global aggdir
        global temppath
        global datapath
        choice = None
        choice = str(vari.get())
       

        if choice == "1":
            ptype = 'CAHPS'
            Data = ra.get_sqlite_data(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\Schedule\CAHPS16.db",
                                'Projects')
            basedir = const.BASEDIR
            temppath = const.AGGREGATES_STORE
            datapath = const.DATA_FILES
        if choice == "2":
            ptype = 'MCAHPS'
            Data = ra.get_sqlite_data(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics Medicare CAHPS\Schedule\MCAHPS17.db",
                             'Projects')
            basedir = mconst.BASEDIR
            temppath = mconst.AGGREGATES_STORE
            datapath = mconst.DATA_FILES

        # Activate directory creation button
        self.create_agg_button.config(highlightbackground="dark gray", highlightthickness=1, state="normal")

    # CREATE DIRECTORY FOR AGGREGATE FILES
    def create_agg_directory(self):
        global aggregate_id
        global aggdir

        aggregate_id = self.aggregate_field.get()
        aggdir = temppath + "\\" + aggregate_id
        print "Aggregate project ID: ", aggregate_id

        # Create storage directory
        if os.path.isdir(aggdir) is False:
            ra.aggdirectory(temppath, aggregate_id)
            print "Directory for aggregate project ID " + aggregate_id + " created."
        else:
            print "Directory already exists, copying required files..."
        
        # Copy RBD Merge tool and RBD Template files to the Aggregates directory
        ra.cp(const.RBD_MERGE, temppath + "\\" + aggregate_id + r"\Templates\2017 RBD Merge.xls")
        ra.cp(const.RBD_TEMPLATE, temppath + "\\" + aggregate_id + r"\Templates\2017 RBD Template.xls")

        # Make file confirmation button active
        self.confirm_button.config(highlightbackground="dark gray", highlightthickness=1, state="normal")

    # CONFIRM THE FILES TO BE COPIED AND USED
    def confirm_files(self):
        global files
        global entries
        files = []
        entries = []

        # FILE 01
        if len(self.id_field_1.get()) == 0:
            pass
        else:
            files.append(self.id_field_1.get())
        # FILE 02
        if len(self.id_field_2.get()) == 0:
            pass
        else:
            files.append(self.id_field_2.get())
        # FILE 03
        if len(self.id_field_3.get()) == 0:
            pass
        else:
            files.append(self.id_field_3.get())
        # FILE 04
        if len(self.id_field_4.get()) == 0:
            pass
        else:
            files.append(self.id_field_4.get())
        # FILE 05
        if len(self.id_field_5.get()) == 0:
            pass
        else:
            files.append(self.id_field_5.get())
        # FILE 06
        if len(self.id_field_6.get()) == 0:
            pass
        else:
            files.append(self.id_field_6.get())
        # FILE 07
        if len(self.id_field_7.get()) == 0:
            pass
        else:
            files.append(self.id_field_7.get())
        # FILE 08
        if len(self.id_field_8.get()) == 0:
            pass
        else:
            files.append(self.id_field_8.get())
        # FILE 09
        if len(self.id_field_9.get()) == 0:
            pass
        else:
            files.append(self.id_field_9.get())
        # FILE 10
        if len(self.id_field_10.get()) == 0:
            pass
        else:
            files.append(self.id_field_10.get())
        # FILE 11
        if len(self.id_field_11.get()) == 0:
            pass
        else:
            files.append(self.id_field_11.get())
        # FILE 12
        if len(self.id_field_12.get()) == 0:
            pass
        else:
            files.append(self.id_field_12.get())
        # FILE 13
        if len(self.id_field_13.get()) == 0:
            pass
        else:
            files.append(self.id_field_13.get())
        # FILE 14
        if len(self.id_field_14.get()) == 0:
            pass
        else:
            files.append(self.id_field_14.get())
        # FILE 15
        if len(self.id_field_15.get()) == 0:
            pass
        else:
            files.append(self.id_field_15.get())
        # FILE 16
        if len(self.id_field_16.get()) == 0:
            pass
        else:
            files.append(self.id_field_16.get())
        # FILE 17
        if len(self.id_field_17.get()) == 0:
            pass
        else:
            files.append(self.id_field_17.get())
        # FILE 18
        if len(self.id_field_18.get()) == 0:
            pass
        else:
            files.append(self.id_field_18.get())
        # FILE 19
        if len(self.id_field_19.get()) == 0:
            pass
        else:
            files.append(self.id_field_19.get())
        # FILE 20
        if len(self.id_field_20.get()) == 0:
            pass
        else:
            files.append(self.id_field_20.get())
        
        print "Text data files will be copied for the following projects:"
        for item in files:
            pull_data = datapath % item
            print ">> " + item
            try:
                assert os.path.isfile(pull_data) # Ensure the file actually exists
            except IOError, AssertionError:
                print "The file " + pull_data + "was not found."

            # Copy data files to Aggregate directory
            copydata = temppath + "\\" + aggregate_id + r"\Files\%s.txt" % item
            ra.cp(pull_data, copydata)
            print "Copied text data for Project ID #" + item + " to Aggregate directory"
        
        # Activate aggregate creation button
        self.create_button.config(highlightbackground="dark gray", highlightthickness=1, state="normal")

    # OUTPUT TEXT BOX 
    def text_ui(self):
        self.output_text = tk.Text(self.main_frame,
                                height=16,
                                width=99,
                                relief="flat",
                                font="Calibri 10",
                                fg="#CCCCCC",
                                bg="#424242")
        self.output_text.config(highlightbackground="dark gray", highlightthickness=0.5)
        self.output_text.place(x=47, y=307)
        sys.stdout = ra.StdoutRedirector(self.output_text)
        sys.stderr = ra.StdoutRedirector(self.output_text)
    
    
    
    # CLEAR TEXT
    def clear_text(self):
        '''
        This function allows the user to clear the text fields and start over.
        '''
        self.id_field_1.delete(0, tk.END)
        self.id_field_2.delete(0, tk.END)
        self.id_field_3.delete(0, tk.END)
        self.id_field_4.delete(0, tk.END)
        self.id_field_5.delete(0, tk.END)
        self.id_field_6.delete(0, tk.END)
        self.id_field_7.delete(0, tk.END)
        self.id_field_8.delete(0, tk.END)
        self.id_field_9.delete(0, tk.END)
        self.id_field_10.delete(0, tk.END)
        self.id_field_11.delete(0, tk.END)
        self.id_field_12.delete(0, tk.END)
        self.id_field_13.delete(0, tk.END)
        self.id_field_14.delete(0, tk.END)
        self.id_field_15.delete(0, tk.END)
        self.id_field_16.delete(0, tk.END)
        self.id_field_17.delete(0, tk.END)
        self.id_field_18.delete(0, tk.END)
        self.id_field_19.delete(0, tk.END)
        self.id_field_20.delete(0, tk.END)
    
    # CLEAR ALL TEXT
    def clear_all(self):
        '''
        This function allows the user to clear the text fields + the aggregate field and start over.
        '''
        self.aggregate_field.delete(0, tk.END)
        self.id_field_1.delete(0, tk.END)
        self.id_field_2.delete(0, tk.END)
        self.id_field_3.delete(0, tk.END)
        self.id_field_4.delete(0, tk.END)
        self.id_field_5.delete(0, tk.END)
        self.id_field_6.delete(0, tk.END)
        self.id_field_7.delete(0, tk.END)
        self.id_field_8.delete(0, tk.END)
        self.id_field_9.delete(0, tk.END)
        self.id_field_10.delete(0, tk.END)
        self.id_field_11.delete(0, tk.END)
        self.id_field_12.delete(0, tk.END)
        self.id_field_13.delete(0, tk.END)
        self.id_field_14.delete(0, tk.END)
        self.id_field_15.delete(0, tk.END)
        self.id_field_16.delete(0, tk.END)
        self.id_field_17.delete(0, tk.END)
        self.id_field_18.delete(0, tk.END)
        self.id_field_19.delete(0, tk.END)
        self.id_field_20.delete(0, tk.END)
    
    # CREATE THE SINGLE AGGREGATE DATA FILE
    def create_aggregate(self, files):
        aggregate_path = temppath + "\\" + aggregate_id + r"\%s.txt" % aggregate_id # Path to aggregate file
        aggregate_log = open(aggregate_path, 'w')

        # Write initial file to new aggregate filename
        first_path = temppath + "\\" + aggregate_id + r'\Files' + r'\%s.txt' % files[0]
        start = open(first_path)
        for i, line in enumerate(start):
                line.strip()
                aggregate_log.write(line)
        start.close()
        aggregate_log.close()
        print "Wrote " + files[0] + '.txt to new aggregate file, ' + aggregate_id + '.txt'

        # Make iterable and skip first file
        iterfiles = iter(files)
        next(iterfiles)

        # Remove headers and combine remaining files
        for item in iterfiles:
            path = temppath + "\\" + aggregate_id + r'\Files' + r'\%s.txt' % item
            piece = open(path)
            aggregate_log = open(aggregate_path, 'a')
            for i, line in enumerate(piece):
                if i > 0:
                    line.strip()
                    aggregate_log.write(line)
            piece.close()
            aggregate_log.close()
            print "Wrote " + item + '.txt to new aggregate file, ' + aggregate_id + '.txt'
        
        # Copy new file to Data Files directory
        ra.cp(aggregate_path, datapath % aggregate_id)
        print "Aggregate file with Project ID " + aggregate_id + " was copied to the Data Files folder."

        



# if __name__ == '__main__':
#     main()
