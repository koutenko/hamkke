#!/usr/bin/env python
"""
Module Name
===========
 $ updatedb.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Shell Syntax
============
 N/A

Description
===========
This module allows one to update the CAHPS SQLite DB.
"""

import sys
import os
import Tkinter as tk
import hamkke as ra
import time
import sqlite3
from sqlite3 import Error
from urlparse import urljoin
from urllib import pathname2url
from subprocess import check_call, CalledProcessError, STDOUT
from pyPdf import PdfFileReader, PdfFileWriter
from wkhtmltopdf import WKhtmlToPdf

class UpdateDB(tk.Frame):
    def __init__(self, frame):
        '''
        This module allows python to update the database and then run a simple
        test to confirm accurate entry.
        '''
        # Initialise the base Tkinter frame
        tk.Frame.__init__(self)

        # Pull data from the 201X CAHPS schedule
        Data = ra.get_sqlite_data(r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Schedule\CAHPS18.db",
                                    'Projects')

        # Create a parent frame to display the module
        top_frame = frame

        # Process to clear anything hanging around before beginning anew
        for widget in top_frame.winfo_children():
            widget.destroy()
        
        # Variable definitions for the Entry boxes
        self.idvar = tk.StringVar()
        self.surveyvar = tk.StringVar()
        self.planvar = tk.StringVar()
        self.q1var = tk.StringVar()
        self.productvar = tk.StringVar()
        self.error_msg = tk.StringVar()
        self.outputvar = tk.StringVar()
        self.checked_surveyvar = tk.StringVar()
        self.checked_planvar = tk.StringVar()
        self.checked_q1var = tk.StringVar()
        self.checked_productvar = tk.StringVar()

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
                            text="- S Q L I T E 3  C A H P S  D A T A B A S E  M A N A G E R -",
                            font="Calibri 18 bold",
                            fg="#E8B93B",
                            bg="#424242")
        mod_title_label.place(x=400, y=12, anchor="center")
        # MODULE INSTRUCTIONS
        instructions_label = tk.Label(self.main_frame, 
                               text="This module allows python to add to and update the SQLite3 CAHPS database remotely.",
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
        #self.v.set("1") # initalise

        self.r1 = tk.Radiobutton(self.main_frame, 
                            text="New Entry", 
                            variable=vari, 
                            value=1, 
                            indicatoron=0,
                            relief="groove",
                            font="Calibri 11",
                            fg="#CCCCCC",
                            bg="#424242",
                            width=10,
                            command=lambda: self.select())
        self.r2 = tk.Radiobutton(self.main_frame,
                            text="Update Entry",
                            variable=vari,
                            value=2,
                            indicatoron=0,
                            relief="groove",
                            font="Calibri 11",
                            fg="#CCCCCC",
                            bg="#424242",
                            width=10,
                            command=lambda: self.select())

        self.r1.place(x=100, y=150)
        self.r2.place(x=100, y=180)

        # -------------- I N F O  F I E L D S --------------
        # PROJECT ID
        self.id_label = tk.Label(self.main_frame, text="Project ID: ", font="Calibri 11 bold", fg="#CCCCCC", bg="#424242")
        self.id_label.place(x=220, y=120)
        
        self.id_field = tk.Entry(self.main_frame, validate="key", relief="flat", textvariable=self.idvar)
        self.id_field.place(x=325, y=120)
        self.id_field.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # SURVEY TYPE
        survey_label = tk.Label(self.main_frame, text="Survey Type: ", font="Calibri 11", fg="#CCCCCC", bg="#424242")
        survey_label.place(x=220, y=150)
        
        self.survey_field = tk.Entry(self.main_frame, relief="flat", textvariable=self.surveyvar)
        self.survey_field.place(x=325, y=150)
        self.survey_field.config(width=10, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # PLAN NAME
        plan_label = tk.Label(self.main_frame, text="Plan Name: ", font="Calibri 11", fg="#CCCCCC", bg="#424242")
        plan_label.place(x=220, y=180)
        
        self.plan_field = tk.Entry(self.main_frame, relief="flat", textvariable=self.planvar)
        self.plan_field.place(x=325, y=180)
        self.plan_field.config(width=40, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # Q1 NAME
        q1_label = tk.Label(self.main_frame, text="Q1 Name: ", font="Calibri 11", fg="#CCCCCC", bg="#424242")
        q1_label.place(x=220, y=210)
        
        self.q1_field = tk.Entry(self.main_frame, relief="flat", textvariable=self.q1var)
        self.q1_field.place(x=325, y=210)
        self.q1_field.config(width=40, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

         # PRODUCT LINE
        product_label = tk.Label(self.main_frame, text="Product Line: ", font="Calibri 11", fg="#CCCCCC", bg="#424242")
        product_label.place(x=220, y=240)
        
        self.product_field = tk.Entry(self.main_frame, relief="flat", textvariable=self.productvar)
        self.product_field.place(x=325, y=240)
        self.product_field.config(width=17, font="Calibri 11", highlightbackground="dark gray", highlightthickness=0.5)

        # -------------- B U T T O N S --------------
        # SEARCH BUTTON
        self.search_button = tk.Button(self.main_frame,
                                  text="Send",
                                  command=lambda: self.do_task(choice),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=7)
        self.search_button.place(x=650, y=150)
        self.search_button.config(highlightbackground="dark gray", highlightthickness=1, state="disabled")

        # CLEAR BUTTON
        self.clear_button = tk.Button(self.main_frame,
                                  text="Clear",
                                  command=lambda: self.clear_text(),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=7)
        self.clear_button.place(x=650, y=185)
        self.clear_button.config(highlightbackground="dark gray", highlightthickness=1)

        # CHECK BUTTON
        self.check_button = tk.Button(self.main_frame,
                                  text="Check Database",
                                  command=lambda: self.do_task("3"),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=14)
        self.check_button.place(x=83, y=265)
        self.check_button.config(highlightbackground="dark gray", highlightthickness=1)

        # FILL BUTTON
        self.fill_button = tk.Button(self.main_frame,
                                  text="Fill Fields",
                                  command=lambda: self.fill_fields(),
                                  relief="groove",
                                  font="Calibri 11",
                                  fg="#CCCCCC",
                                  bg="#424242",
                                  width=14)
        self.fill_button.place(x=603, y=265)
        self.fill_button.config(highlightbackground="dark gray", highlightthickness=1, state="disabled")

        # # RUN PREPWORK BUTTON
        # self.run_button = tk.Button(self.main_frame,
        #                           text="Run Project",
        #                           command=lambda: self.run_task_check(),
        #                           relief="groove",
        #                           font="Calibri 11",
        #                           width=34,
        #                           fg="#CCCCCC",
        #                           bg="#424242")
        # self.run_button.place(x=505, y=185)
        # self.run_button.config(highlightbackground="dark gray", highlightthickness=1, state="disabled")

    def text_ui(self):
        # OUTPUT TEXT BOX
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
    
    def select(self):
        global choice
        choice = None
        choice = str(vari.get())
        self.search_button.config(highlightbackground="dark gray", highlightthickness=1, state="normal")
    
    # CLEAR TEXT
    def clear_text(self):
        '''
        This function allows the user to clear the text fields and start over.
        '''
        self.id_field.config(state="normal")
        self.id_field.delete(0, tk.END)

        self.survey_field.config(state="normal")
        self.survey_field.delete(0, tk.END)

        self.plan_field.config(state="normal")
        self.plan_field.delete(0, tk.END)

        self.q1_field.config(state="normal")
        self.q1_field.delete(0, tk.END)

        self.product_field.config(state="normal")
        self.product_field.delete(0, tk.END)

        self.error_msg.set("")
        self.id_label.config(fg="#CCCCCC",
                              font="Calibri 11 bold")

    def set_new(self, conn, task):
        # idvar = self.id_field.get()
        # surveyvar = self.survey_field.get()
        # planvar = self.plan_field.get()
        # q1var = self.q1_field.get()
        # productvar = self.product_field.get()
        sql = ''' INSERT INTO Projects(ProjectID, SurveyType, PlanName, Q1Name, ProductLine)
                    VALUES (?, ?, ?, ?, ?);'''#, (idvar, surveyvar, planvar, q1var, productvar)
        
        curr = conn.cursor()

        curr.execute('SELECT * FROM Projects WHERE ProjectID = ?', (idvar,))
        entry = curr.fetchone()

        if entry is None:
            curr.execute(sql, task)
            conn.commit()
            print "INSERT INTO Projects(ProjectID, SurveyType, PlanName, Q1Name, ProductLine"
            print "VALUES (%s, " % task[0] + "%s, " % task[1] + "%s, " % task[2] + "%s, " % task[3] + "%s)" % task[4]
            print "New entry added.\n"
        else:
            print "An entry with this Project ID already exists. If you need to update the information, run an Update.\n"
        
    def set_update(self, conn, task):
        # idvar = self.id_field.get()
        # surveyvar = self.survey_field.get()
        # planvar = self.plan_field.get()
        # q1var = self.q1_field.get()
        # productvar = self.product_field.get()
        sql = ''' UPDATE Projects
                SET SurveyType = ? ,
                    PlanName = ? ,
                    Q1Name = ? ,
                    ProductLine = ?
                WHERE ProjectID = ?'''#, (surveyvar, planvar, q1var, productvar, idvar)
        
        curr = conn.cursor()
        curr.execute(sql, task)
        conn.commit()

        print "UPDATE Projects"
        print "SET SurveyType = ", task[0]
        print "> PlanName = ", task[1]
        print "> Q1Name = ", task[2]
        print "> ProductLine = ", task[3]
        print "WHERE ProjectID = ", task[4]
        print ""
        
    def check_database(self, conn, task):
        # idvar = self.id_field.get()
        # surveyvar = self.survey_field.get()
        # planvar = self.plan_field.get()
        # q1var = self.q1_field.get()
        # productvar = self.product_field.get()
        global checked_surveyvar
        global checked_planvar
        global checked_q1var
        global checked_productvar

        survey = ""
        plan = ""
        q1 = ""
        product = ""

        sql = ''' SELECT ProjectID, SurveyType, PlanName, Q1Name, ProductLine
                  FROM Projects
                  WHERE ProjectID = ?'''
        
        curr = conn.cursor()
        curr.execute(sql, (task,))
        
        for row in curr:
            survey = row[1]
            plan = row[2]
            q1 = row[3]
            product = row[4]

            print "SELECT * FROM Projects WHERE Project ID = ", row[0]
            print "> Project ID: ", row[0]
            print "> Survey Type: ", row[1]
            print "> Plan Name: ", row[2]
            print "> Q1 Name: ", row[3]
            print "> Product Line: ", row[4]
            print ""
        
        checked_surveyvar = survey
        checked_planvar = plan
        checked_q1var = q1
        checked_productvar = product
        
        # Unlock option to fill in fields
        self.fill_button.config(highlightbackground="dark gray", highlightthickness=1, state="normal")
        
    def create_connection(self, db_file):
        try:
            conn = sqlite3.connect(db_file)
            return conn
        except Error as e:
            print(e)
        
        return None

    def do_task(self, v):
        database= r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Schedule\CAHPS18.db"
        #database = r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\Schedule\CAHPS16.db"
        #database = r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\Schedule\CAHPS16.db"
        global idvar
        global surveyvar
        global planvar
        global q1var
        global productvar
        
        idvar = self.id_field.get()
        surveyvar = self.survey_field.get()
        planvar = self.plan_field.get()
        q1var = self.q1_field.get()
        productvar = self.product_field.get()

        in_params = (idvar, surveyvar, planvar, q1var, productvar)
        up_params = (surveyvar, planvar, q1var, productvar, idvar)        

        # create a database connection
        conn = self.create_connection(database)
        with conn:
            if v == "1":
                self.set_new(conn, in_params)
            if v == "2":
                self.set_update(conn, up_params)
            if v == "3":
                try:
                    self.check_database(conn, idvar)
                except Error as e:
                    print(e)
            # insert_task(conn, (projectID, survey_type, plan_name, q1_name, product_line))

    # Plug values into fields for updates
    def fill_fields(self):
        '''
        This function allows the user to input the Project ID or the
        Contract Number and the other fields will propagate automatically
        after the database is checked. If the entry exists it will plug the data
        into the fields, saving time.
        '''
        self.surveyvar.set(checked_surveyvar)
        self.planvar.set(checked_planvar)
        self.q1var.set(checked_q1var)
        self.productvar.set(checked_productvar)
    

# if __name__ == '__main__':
#     main()
