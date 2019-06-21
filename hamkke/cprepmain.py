#!/usr/bin/env python
"""
Module Name
===========
 $ cprepmain.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Description
===========
The module contains the scripting necessary to run prep work functions on CAHPS projects.
It is called within the main program from the top bar menu. It must be run as a prerequisite
to running the Reports module.

The GUI class is instantiated when this module is run.

This module calls various "get" methods in order to fill variables. This ensures paths are cleanly
and reliably captured before running any Syntax or Script operations.

These scripts do the exact same thing as the cahpsprep.py Command Line script.
"""

import os
import sys
import shlex
import glob
import time
import re
import pdb
import shutil
import threading
from datetime import datetime as dt
from datetime import timedelta
from subprocess import Popen, PIPE
from multiprocessing import Queue, Process
from jinja2 import Environment, PackageLoader
import Tkinter as tk
import ttk
import tkMessageBox
import hamkke as ra
from hamkke.cahps import constants as const


# ---------------- START FUNCTION ----------------
def get_gui(frame):
    '''
    This function calls the cPrepGUI() class and
    assigns it to a global variable.
    '''
    global ui
    ui = ra.cPrepGUI(frame)


# ------- BEGIN VARIABLE ASSIGNMENT METHODS -------
def get_basedir():
    '''
    Get base directory path.
    '''
    basedir = const.BASEDIR

    return basedir

def get_contractdir(contract):
    '''
    Get contract directory path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']

    return contractdir

def get_mrstemp(contract):
    '''
    Get the correct .mrs template according to Survey Type.
    '''
    if contract['SurveyType'] == 'CAS':
        mrstemp = const.CORE_MRS_CAS
    if contract['SurveyType'] == 'MAS':
        mrstemp = const.CORE_MRS_MAS
    if contract['SurveyType'] == 'MCS':
        mrstemp = const.CORE_MRS_MCS
    if contract['SurveyType'] == 'MCS CCC':
        mrstemp = const.CORE_MRS_MCS_CCC

    return mrstemp

def get_spstemp(contract):
    '''
    Get the correct sps template according to Survey Type.
    '''
    if contract['SurveyType'] == 'CAS':
        spstemp = const.CORE_SPS_CAS
    if contract['SurveyType'] == 'MAS':
        spstemp = const.CORE_SPS_MAS
    if contract['SurveyType'] == 'MCS':
        spstemp = const.CORE_SPS_MCS
    if contract['SurveyType'] == 'MCS CCC':
        spstemp = const.CORE_SPS_MCS_CCC

    return spstemp

def get_spj(contract):
    '''
    Get path for spj file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']

    spjtemp = contractdir + '\\Syntax\\' + '%s' % contract['ProjectID'] + 'syntax.spj'

    return spjtemp

def get_spv(contract):
    '''
    Get path for spv file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    spvtemp = contractdir + '\\Syntax\\' + '%s' % contract['ProjectID'] + '.spv'

    return spvtemp

def get_bannerpath(contract):
    '''
    Get the banner path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']

    bannerpath = contractdir + '\\Banners'

    return bannerpath

def get_banner_html(contract):
    '''
    Get the banner HTML path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    if contract['SurveyType'] == 'MCS CCC':
        bannerhtml = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners CCC POP.htm'
    else:
        bannerhtml = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners.htm'

    return bannerhtml

def get_mcs_ccc_banners_html(contract):
    '''
    Get the banner HTML path for MCS CCC banners.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    gen_banner_html = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners GEN POP.htm'
    ccc_banner_html = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners CCC POP.htm'

    return gen_banner_html, ccc_banner_html

def get_banner_pdf(contract):
    '''
    Get the banner PDF path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    bannerpdf = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners.pdf'

    return bannerpdf

def get_mcs_ccc_banners_pdf(contract):
    '''
    Get the banner PDF path for MCS CCC banners.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    gen_banner_pdf = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners GEN POP.pdf'
    ccc_banner_pdf = contractdir + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners CCC POP.pdf'

    return gen_banner_pdf, ccc_banner_pdf


def get_text_data(contract):
    '''
    Get the text data path.
    '''
    basedir = const.BASEDIR
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    textdata = contractdir + '\\Data\\%s.txt' % contract['ProjectID']    

    return textdata

def get_RBD(contract):
    '''
    Get the RBD path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    rbd_path = contractdir + '\\Data\\%s' % contract['ProjectID'] + 'RBD.xls'

    return rbd_path

def get_RBD_A(contract):
    '''
    Get the RBD_A path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    rbd_a_path = contractdir + '\\Data\\%s' % contract['ProjectID'] + 'RBD_A.xls'

    return rbd_a_path

def get_temp_path(contract):
    '''
    Get the temporary directory (clones folder) path for 
    the template files.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    temp_path = contractdir + '\\clones'

    return temp_path

def prep_clones(contract, path):
    '''
    Get the appropriate Excel and Word template file
    and copy them to the temporary clone directory.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    temp_path = contractdir + '\\clones'

    excel_temp = const.EXCEL_TEMPLATE % contract['SurveyType']
    ra.cp(excel_temp, temp_path)

    word_temp = const.WORD_TEMPLATE % contract['SurveyType']
    ra.cp(word_temp, temp_path)

    return 0

def del_clones(contractdir):
    '''
    Delete the excess clones folder.
    '''
    # Delete clones directory
    existing = os.listdir(contractdir)
    for folder in existing:
        if folder == "clones":
            shutil.rmtree('%s' % contractdir + '\\clones')

    return 0

# --------------- GET PROJECT DATA ----------------
def start_prep(projectid):
    # Data = ra.get_project_info(const.FIELDS, const.CONTRACTLIST,
    #                         delimiter="|", headers=True)

    # Pull data from the 2016 CAHPS schedule.
    Data = ra.get_sqlite_data(r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Schedule\CAHPS18.db",
                                'Projects')
    
    # Assign paths to SPSS Stats and mrScript Command Line.
    global stats_path
    stats_path = const.STATS
    global script_path
    script_path = const.MRSCRIPTCL
    global wkpdf_path
    wkpdf_path = const.WKPDF

    args = projectid
    contract = args

    allstart = dt.now()
    alltimes = []

    _ = Data[contract]           

    runtime = prep_work(Data[contract], debug=False, noarchive=False)

    alltimes.append(runtime)
    allend = dt.now()

    ''' if len(alltimes) != 0:
        print "Average contract running time: ", \
                sum(alltimes, timedelta()) / len(alltimes) '''

    return 0

# Module to call and execute the Syntax and Script modules, generating syntax and banners.
def prep_work(project, debug=False, noarchive=False):
    '''
    This module defines the variables and calls other functions to run
    the different processes necessary to generate the files that will
    be used to create the final report.
    '''
    # Define global variables
    global contract
    global basedir
    global contractdir
    global mrstemp
    global spstemp
    global temp_path
    global mrscopy
    global spscopy
    global spjtemp
    global spvtemp
    global bannerpath
    global bannerhtml
    global bannerpdf
    global textdata
    #global rbd_path
    #global rbda_path

    start = dt.now()  # Start timing
    contract = project

    gui_toggle = True # denote that we are using the GUI version

    # Start Progress bar
    ui.progress_bar()

    # Get the Reporting path
    basedir = get_basedir()
    print "\n> REPORTING PATH: \n>> %s" % basedir

    # Get the contract path
    contractdir = get_contractdir(contract)
    print "> CONTRACT PATH: \n>> %s" % contractdir

    # Make the contract's reporting folder
    if not noarchive:
        ra.makedirectory(basedir, contract['ProjectID'], gui_toggle)

    # Decide what templates to use
    mrstemp = get_mrstemp(contract)
    spstemp = get_spstemp(contract)

    # Clone templates in the contract folder to use
    print "\nCloning templates...",
    temp_path = get_temp_path(contract)
    mrscopy = ra.copy_mrs_template(mrstemp, temp_path, contractdir)
    spscopy = ra.copy_sps_template(spstemp, temp_path, contractdir)
    check_clones(mrscopy, spscopy)
    #check_clones(mrscopy)
    
    # Get the paths
    spjtemp = get_spj(contract)
    spvtemp = get_spv(contract)
    bannerpath = get_bannerpath(contract)
    bannerhtml = get_banner_html(contract)
    bannerpdf = get_banner_pdf(contract)
    textdata = get_text_data(contract)
    #rbd_path = get_RBD(contract)
    #rbda_path = get_RBD_A(contract)

    # Run the Syntax
    check_syntax()
    
    # Run the Script
    check_script()
    
    # Create the PDF Banners
    check_banners()

    # Copy the Data
    check_copies()

    # Copy the Templates
    #check_templates()

    # Delete the Clones
    delete_clones()

    end = dt.now()  # Stop timing
    runtime = end - start

    '''if ra.compare_timestamps(textdata, rbd_path):
            print "\nNote:"
            print "The timestamps for the raw text data file and the RBD are more"
            print "than 3 minutes apart.  This could indicate that one of those"
            print "components did not export correctly.  Consider double-checking.\n"
    print "Process completed in %s.\n" % runtime
    return runtime'''

def check_module(projectid):
    '''
    Check the process status of the overall module
    and update the labels.
    '''
    ui.module_check.config(fg="#B68C14")
    ui.module_check_text.config(fg="#B68C14")
    ui.update_idletasks()
    
    pass_id = projectid
    status = start_prep(pass_id)

    time.sleep(0.1)
    ui.progress.update()

    if status == 0:
        # Update Module Check status
        ui.module_check.config(fg="#60E83B")
        ui.module_check_text.config(fg="#60E83B")
        ui.update_idletasks()
    else:
        # Update Module Check status
        ui.module_check.config(fg="#E8483B")
        ui.module_check_text.config(fg="#E8483B")
        ui.mod_passfail.set(ui.failpass)
        ui.update_idletasks()
        
#def check_clones(mrscopy, spscopy):
def check_clones(mrscopy, spscopy):
    '''
    Check the process status of the first template
    cloning process and update the labels.
    '''
    # Update Syntax Templates Check status
    ui.syn_template_check.config(fg="#B68C14")
    ui.syn_template_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    #if os.path.isfile(mrscopy) and os.path.isfile(spscopy):
    if os.path.isfile(mrscopy):
        # Update Syntax Templates Check status
        ui.syn_template_check.config(fg="#60E83B")
        ui.syn_template_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Syntax Templates Check status
        ui.syn_template_check.config(fg="#E8483B")
        ui.syn_template_check_text.config(fg="#E8483B")
        ui.s_temp_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

def check_syntax():
    '''
    Check the process status of the Syntax running
    and update the labels.
    '''
    # Update Run Syntax Check status
    ui.data_check.config(fg="#B68C14")
    ui.data_check_text.config(fg="#B68C14")
    ui.run_syn_check.config(fg="#B68C14")
    ui.run_syn_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    # Run the syntax
    symbols = {"@ProjectID": contract['ProjectID']}
    syntax = ra.Syntax(projectid=contract['ProjectID'], spj=spjtemp, sps=spscopy, spv=spvtemp,
                    stype='core', symbols=symbols, stats=stats_path, debug=False)
    status = syntax.run(debug=False)

    if status == 0:
        # Update Run Syntax Check status
        ui.data_check.config(fg="#60E83B")
        ui.data_check_text.config(fg="#60E83B")
        ui.run_syn_check.config(fg="#60E83B")
        ui.run_syn_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Run Syntax Check status
        ui.data_check.config(fg="#E8483B")
        ui.data_check_text.config(fg="#E8483B")
        ui.run_syn_check.config(fg="#E8483B")
        ui.run_syn_check_text.config(fg="#E8483B")
        ui.syn_passfail.set(ui.failpass)
        ui.data_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

def check_script():
    '''
    Check the process status of the Scripts running
    and update the labels.
    '''
    # Update Run Script Check status
    ui.script_check.config(fg="#B68C14")
    ui.script_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    # Run the script
    constants = {'ProjectID': '%s' % contract['ProjectID'],
                'PlanName': '%s' % contract['PlanName'].replace('&', '&amp;'),
                'Q1Name': '%s' % contract['Q1Name'].replace('&', '&amp;'),
                'SurveyType': '%s' % contract['SurveyType']}
    script = ra.SRScript(mrscopy, stype='core', constants=constants, script=script_path,
                            bannerpath=bannerpath)
    status = script.run(debug=False)

    if status == 0:
        # Update Run Script Check status
        ui.script_check.config(fg="#60E83B")
        ui.script_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Run Script Check status
        ui.script_check.config(fg="#E8483B")
        ui.script_check_text.config(fg="#E8483B")
        ui.script_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

def check_banners():
    '''
    Check the process status for creating the banner files
    and update the labels.
    '''
    # Update Banners Check status
    ui.banner_check.config(fg="#B68C14")
    ui.banner_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    # Create PDF banner tables from HTML banner file.
    if contract['SurveyType'] == 'MCS CCC':
        gen_banner_html, ccc_banner_html = get_mcs_ccc_banners_html(contract)
        gen_banner_pdf, ccc_banner_pdf = get_mcs_ccc_banners_pdf(contract)

        ra.pdf_html_file(gen_banner_html, dst=gen_banner_pdf, wkpdf_path=wkpdf_path, overwrite=1, debug=False)
        ra.pdf_html_file(ccc_banner_html, dst=ccc_banner_pdf, wkpdf_path=wkpdf_path, overwrite=1, debug=False)
    else:
        ra.pdf_html_file(bannerhtml, dst=bannerpdf, wkpdf_path=wkpdf_path, overwrite=1, debug=False)
    
    if contract['SurveyType'] == 'MCS CCC':
        if os.path.isfile(gen_banner_pdf) and os.path.isfile(ccc_banner_pdf):
            # Update Banners Check status
            ui.banner_check.config(fg="#60E83B")
            ui.banner_check_text.config(fg="#60E83B")
            time.sleep(0.1)
            ui.progress.step()
            ui.progress.update()
        else:
            # Update Banners Check status
            ui.banner_check.config(fg="#E8483B")
            ui.banner_check_text.config(fg="#E8483B")
            ui.banner_passfail.set(ui.failpass)
            time.sleep(0.1)
            ui.progress.step()
            ui.progress.update()
    else:
        if os.path.isfile(bannerpdf):
            # Update Banners Check status
            ui.banner_check.config(fg="#60E83B")
            ui.banner_check_text.config(fg="#60E83B")
            time.sleep(0.1)
            ui.progress.step()
            ui.progress.update()
        else:
            # Update Banners Check status
            ui.banner_check.config(fg="#E8483B")
            ui.banner_check_text.config(fg="#E8483B")
            ui.banner_passfail.set(ui.failpass)
            time.sleep(0.1)
            ui.progress.step()
            ui.progress.update()

def check_copies():
    '''
    Check the process status for copying the data and 
    RBD files and update the labels.
    '''
    # Update Copy Check status
    ui.copy_check.config(fg="#B68C14")
    ui.copy_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    status = 0

    # Copy text data and RBD to Data folder
    print ">> Creating Data folder copies...",
    try:
        ra.cp(r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Data Files\%s.txt" % contract['ProjectID'],
            textdata)
        ''' ra.cp(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\RBD\%sRBD.xls" % contract['ProjectID'],
            rbd_path) 
        if contract['SurveyType'] == 'MCS CCC':
            ra.cp(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\RBD\%sRBD_A.xls" %
                contract['ProjectID'], rbda_path)'''
    except IOError:
        pass
        status = 1
    print "Success!"
    
    # Copy text data and RBD to clones folder
    print ">> Cloning text data...",
    try:
        ra.cp(r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Data Files\%s.txt" % contract['ProjectID'],
            temp_path)
        '''ra.cp(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\RBD\%sRBD.xls" % contract['ProjectID'],
            temp_path)
        if contract['SurveyType'] == 'MCS CCC':
            ra.cp(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\RBD\%sRBD_A.xls" %
                contract['ProjectID'], temp_path)'''
    except IOError:
        pass
        status = 1
    print "Success!"

    if status == 0:
        # Update Copy Check status
        ui.copy_check.config(fg="#60E83B")
        ui.copy_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Copy Check status
        ui.copy_check.config(fg="#E8483B")
        ui.copy_check_text.config(fg="#E8483B")
        ui.copy_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

def check_templates():
    '''
    Check the process status for copying the report templates
    and update the labels.
    '''
    ui.rep_template_check.config(fg="#B68C14")
    ui.rep_template_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    # Copy final report templates to clones folder
    print ">> Cloning Excel and Word templates...",
    status = prep_clones(contract, temp_path)
    print "Success!"

    if status == 0:
        # Update Template Check status
        ui.rep_template_check.config(fg="#60E83B")
        ui.rep_template_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Template Check status
        ui.rep_template_check.config(fg="#E8483B")
        ui.rep_template_check_text.config(fg="#E8483B")
        ui.rep_temp_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

def delete_clones():
    '''
    Check the progress status for deleting the excess clones directory.
    '''
    ui.del_clones_check.config(fg="#B68C14")
    ui.del_clones_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    # Copy final report templates to clones folder
    print ">> Deleting clones folder...",
    status = del_clones(contractdir)
    print "Success!"

    if status == 0:
        # Update Delete Clones Check status
        ui.del_clones_check.config(fg="#60E83B")
        ui.del_clones_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Delete Clones Check status
        ui.del_clones_check.config(fg="#E8483B")
        ui.del_clones_check_text.config(fg="#E8483B")
        ui.del_clones_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

def copy_banners():
    '''
    Check the process status for copying the banners to the Reporting folder and 
    and update the labels.
    '''
    # Update Banner Check status
    ui.copy_banner.config(fg="#B68C14")
    ui.copy_banner_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    status = 0

    orig_path = r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Banners\%s" % contract['ProjectID']
    reporting_dir = r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Reporting\%s" % contract['ProjectID']
    reporting_path = reporting_dir + '\\Banners\\'
    htm_path = orig_path + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners.htm'
    pdf_path = orig_path + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners.pdf'
    csv1_path = orig_path + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners1.csv'
    csv2_path = orig_path + '\\Banners\\' + '%s' % contract['ProjectID'] + ' Banners2.csv'
    logo_path = r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS\Banners\%s\Banners\logo.png" % contract['ProjectID']
    print htm_path
    print pdf_path
    print csv1_path
    print csv2_path
    print logo_path

    # Copy Banners to Reporting folder
    print ">> Creating Reporting folder copies...",
    try:
        ra.cp(htm_path, reporting_path)
        ra.cp(pdf_path, reporting_path)
        ra.cp(csv1_path, reporting_path)
        ra.cp(csv2_path, reporting_path)
        ra.cp(logo_path, reporting_path)
    except IOError:
        pass
        status = 1
    print "Success!"

    if status == 0:
        # Update Banner Check status
        ui.copy_banner.config(fg="#60E83B")
        ui.copy_banner_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()
    else:
        # Update Copy Banner status
        ui.copy_banner.config(fg="#E8483B")
        ui.copy_banner_text.config(fg="#E8483B")
        ui.cbanner_passfail.set(ui.failpass)
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

    
            
        



