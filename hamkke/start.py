#!/usr/bin/env python
"""
Script Name
===========
 $ start.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Shell Syntax
============


Description
===========
This script is used to run the hamkke Report Automation module with graphics.
"""
import sys
import os
import shutil
import Tkinter as tk
from tkMessageBox import *
import hamkke as ra
from hamkke.cahps import constants as const

class MainApplication(object):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        tk.Frame.minsize(width=600, height=400)
        tk.Frame.maxsize(width=600, height=400)
        tk.Frame.resizable(width=False, height=False)
        tk.Frame.wm_title("hamkke")

        self.menu = ra.Menu(tk.Frame)

        #cahpsmod = False

        # frame = tk.Frame(master)
        # frame.grid()

        # menubar = tk.Menu(master)
        
        # modmenu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="File", menu=modmenu)
        # modmenu.add_command(label="Exit", command=master.quit)
        # master.config(menu=menubar)

        # cahpsmenu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="CAHPS", menu=cahpsmenu)
        self.menu.cahpsmenu.add_command(label="Run Prepwork", command=lambda: self.cahps_prepmod(frame))
        self.menu.cahpsmenu.add_command(label="Generate Reports", command=lambda: self.cahps_reportmod(frame))

        # mcahpsmenu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="MCAHPS", menu=mcahpsmenu)
        self.menu.mcahpsmenu.add_command(label="Run Prepwork", command=lambda: self.mcahps_prepmod(frame))
        self.menu.mcahpsmenu.add_command(label="Generate Reports", command=lambda: self.mcahps_reportmod(frame))

        self.menu.toolsmenu.add_command(label="Create Aggregate Data", command=lambda: self.create_aggregate(frame))
        self.menu.toolsmenu.add_command(label="Update Database", command=lambda: self.update_data(frame))

        # if cahpsmod is False:
        #     self.menu.cahpsmenu.entryconfig(1, state="disabled")
        # else:
        #     self.menu.cahpsmenu.entryconfig(1, state="enabled")
    
    def cahps_prepmod(self, frame):
        '''
        Call prepmod class.
        '''
        cahpsmod = ra.cPrepMain(frame)
        cahpsmod.pre_prep()

        return cahpsmod

    def cahps_reportmod(self, frame):
        '''
        Call reportmod class.
        '''
        cahpsreport = ra.cReportsMain(frame)

        return cahpsreport
    
    def update_data(self, frame):
        '''
        Call db updater class.
        '''
        update_db = ra.UpdateDB(frame)

        return update_db
    



















    def cahps_reportmod(self, frame):
        print "Called"
        frame = frame
        # if (self.cahps_frame.winfo_exists()):
        #     self.cahps_frame.isgridded = False
        #     self.cahps_frame.grid_forget()
        for widget in frame.winfo_children():
            widget.destroy()

        new_frame = tk.Frame(frame)
        new_frame.grid(row=0, column=0)
        new_frame.isgridded = False #Dynamically add "isgridded" attribute.
        mcahps_label = tk.Label(new_frame, text="Welcome to MCAHPS Report Automation!")
        mcahps_label.grid(row=0, column=1)

    # MCAHPS Modules
    def mcahps_prepmod(self, frame):
        print "Called"
        frame = frame
        # if (self.mcahps_frame.winfo_exists()):
        #     self.mcahps_frame.isgridded = False
        #     self.mcahps_frame.grid_forget()
        for widget in frame.winfo_children():
            widget.destroy()

        new_frame = tk.Frame(frame)
        new_frame.grid(row=0, column=0)
        new_frame.isgridded = False #Dynamically add "isgridded" attribute.
        cahps_label = tk.Label(new_frame, text="Welcome to CAHPS Report Automation!")
        cahps_label.grid(row=0, column=1)

    def mcahps_reportmod(self, frame):
        print "Called"
        frame = frame
        # if (self.cahps_frame.winfo_exists()):
        #     self.cahps_frame.isgridded = False
        #     self.cahps_frame.grid_forget()
        for widget in frame.winfo_children():
            widget.destroy()

        new_frame = tk.Frame(frame)
        new_frame.grid(row=0, column=0)
        new_frame.isgridded = False #Dynamically add "isgridded" attribute.
        mcahps_label = tk.Label(new_frame, text="Welcome to MCAHPS Report Automation!")
        mcahps_label.grid(row=0, column=1)

root = tk.Tk()
app = MainApplication(root)
root.mainloop()

# base=tk.Tk()  #this is the main frame
# root=tk.Frame(base)  #Really this is not necessary -- the other widgets could be attached to "base", but I've added it to demonstrate putting a frame in a frame.
# root.grid(row=0,column=0)
# scoreboard=tk.Frame(root)
# scoreboard.grid(row=0,column=0,columnspan=2)

# ###
# #Code to add stuff to scoreboard ...
# # e.g. 
# ###
# scorestuff=tk.Label(scoreboard,text="Here is the scoreboard")
# scorestuff.grid(row=0,column=0)
# #End scoreboard

# #Start cards.
# cards=tk.Frame(root)
# cards.grid(row=1,column=0)
# ###
# # Code to add pitcher and batter cards
# ###
# clabel=tk.Label(cards,text="Stuff to add cards here")
# clabel.grid(row=0,column=0)
# #end cards

# #Offense/Defense frames....
# offense=tk.Frame(root)
# offense.grid(row=1,column=1)
# offense.isgridded=True #Dynamically add "isgridded" attribute.
# offense_label=tk.Label(offense,text="Offense is coolest")
# offense_label.grid(row=0,column=0)

# defense=tk.Frame(root)
# defense.isgridded=False
# defense_label=tk.Label(defense,text="Defense is coolest")
# defense_label.grid(row=0,column=0)

# def switchOffenseDefense():
#     print "Called"
#     if(offense.isgridded):
#         offense.isgridded=False
#         offense.grid_forget()
#         defense.isgridded=True
#         defense.grid(row=1,column=1)
#     else:
#         defense.isgridded=False
#         defense.grid_forget()
#         offense.isgridded=True
#         offense.grid(row=1,column=1)


# switch_button=tk.Button(root,text="Switch",command=switchOffenseDefense)
# switch_button.grid(row=2,column=1)

# root.mainloop()
