#!/usr/bin/env python
"""
Module Name
===========
 $ creportsmain.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Description
===========
This module runs the GUI and scripting to run prep work functions on CAHPS projects.
It is called within the main program from the top bar menu. It must be run as a prerequisite
to running the Reports module.

This module calls various "get" methods in order to fill variables. This ensures paths are cleanly
and reliably captu#E8483B before running any Syntax or Script operations.
"""

import os
import sys
import shlex
import glob
import time
import re
import pdb
import threading
import shutil
from datetime import datetime as dt
from datetime import timedelta
from subprocess import Popen, PIPE
from multiprocessing import Queue, Process
from jinja2 import Environment, PackageLoader
import Tkinter as tk
import ttk
import tkMessageBox
import hamkke as ra
from hamkke import cp
from hamkke.cahps import constants as const
from hamkke.prepcahps import PrepCahps as prep
import hamkke.excel as excel
import hamkke.word as word
import hamkke.cahps.pdfmerge as pdfmerge
import hamkke.errors as errors
from hamkke.projectinfo import get_project_info, get_sqlite_data
from hamkke.utils import merge_report, pdf_html_file

# ---------------- PREP SCRIPT ----------------
def get_gui(frame):
    '''
    This function calls the cReportsGUI() class and
    assigns it to a global variable.
    '''
    global ui
    ui = ra.cReportsGUI(frame)

# ---- BEGIN VARIABLE ASSIGNMENT METHODS ----
def get_temp_path(contract):
    '''
    Get the path to the clones directory.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    temp_path = contractdir + '\\clones'

    return temp_path

# RBD path methods.    
def get_RBD(contract, temp_path):
    '''
    Get the RBD path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    rbd_path = temp_path + '\\%s' % contract['ProjectID'] + 'RBD.xls'

    return rbd_path

def get_RBD_A(contract, temp_path):
    '''
    Get the RBD_A path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    rbd_a_path = temp_path + '\\%s' % contract['ProjectID'] + 'RBD_A.xls'

    return rbd_a_path

# Banner path methods.
def get_banners_1(contract):
    '''
    Get the Banners1 path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    banners_1 = contractdir + '\\Banners\\' + '%s Banners1.csv' % contract['ProjectID']

    return banners_1


def get_banners_2(contract):
    '''
    Get the Banners2 path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    banners_2 = contractdir + '\\Banners\\' + '%s Banners2.csv' % contract['ProjectID']

    return banners_2

# Methods to get the correct Excel files.
def get_excel_temp(contract, temp_path):
    '''
    Get the appropriate Excel template file.
    '''
    excel_temp = temp_path + '\\2017 %s Template Report.xlsx' % contract['SurveyType']

    return excel_temp


def get_excel_report(contract):
    '''
    Get the appropriate Excel report file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    excel_report = contractdir + '\\Final Report\\' + '%s' % contract['ProjectID'] + ' %s' % contract['SurveyType'] + ' Excel Report.xlsx'

    return excel_report

def get_excel_pdf(contract):
    '''
    Get the appropriate Excel PDF file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    excel_pdf = contractdir + '\\Final Report\\' + '%s' % contract['ProjectID'] + ' %s' % contract['SurveyType'] + ' Excel Report.pdf'

    return excel_pdf

# Methods to get the correct Word files.
def get_word_temp(contract, temp_path):
    '''
    Get the appropriate Word template file.
    '''
    word_temp = temp_path + '\\2017 %s Template Report.docx' % contract['SurveyType']

    return word_temp

def get_word_report(contract):
    '''
    Get the appropriate Word report file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    word_report = contractdir + '\\Final Report\\' + '%s' % contract['ProjectID'] + ' %s' % contract['SurveyType'] + ' Word Report.docx'

    return word_report

def get_word_pdf(contract):
    '''
    Get the appropriate Word PDF file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    word_pdf = contractdir + '\\Final Report\\' + '%s' % contract['ProjectID'] + ' %s' % contract['SurveyType'] + ' Word Report.pdf'

    return word_pdf

# Method to get correct Mail Merge files.
def get_merge_temp(contract):
    '''
    Get the appropriate Mail Merge Template file.
    '''
    merge_template = const.MERGE_TEMPLATE % contract['SurveyType']
    ra.cp(merge_template, temp_path)
    merge_temp = temp_path + '\\%s' % contract['SurveyType'] + ' Mail Merge.csv'

    return merge_temp

def get_mail_merge(contract):
    '''
    Get the appropriate Merge Data file.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    mail_merge = contractdir + '\\Final Report\\MergeData.csv'

    return mail_merge

# Get CQ paths
def get_cq_book(contract):
    '''
    Get CQ Book path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    cq_book = contractdir + '\\CQ\\' + '%s CQ Book.xlsm' % contract['ProjectID']

    return cq_book

def get_cq_book_pdf(contract):
    '''
    Get CQ Book PDF path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    cq_book_pdf = contractdir + '\\CQ\\' + '%s CQ Book.pdf' % contract['ProjectID']

    return cq_book_pdf

# Get Final PDF path
def get_final_pdf(contract):
    '''
    Get final PDF path.
    '''
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    final_pdf = contractdir + '\\Final Report\\' + '%s' % contract['ProjectID'] + ' 2017 CAHPS %s Report.pdf' % contract['SurveyType']

    return final_pdf
# ---- END VARIABLE ASSIGNMENT METHODS ----

def start_reports(projectid):
    '''
    Run for selected or run for all.

    The Word report module will only run if the Excel
    report was produced successfully.
    '''
    # Pull data from the 2017 CAHPS schedule.
    Data = ra.get_sqlite_data(r"\\P10-FILESERV-01.sphanalytics.org\PM2017_Peak10\Analytics 2017\Analytics CAHPS\Schedule\CAHPS16.db",
                            'Projects')
    
    args = projectid
    contracts = args

    allstart = dt.now()

    contract = Data[contracts]
    
    # Runs the Excel report module
    xlfail = begin_excel(contract)
    if not xlfail:
        # Runs the Word report module
        wrdfail = begin_word(contract)
    #    if not wrdfail:
    #        # Runs the Merge Components process.
    #        merge_components(Data[contract])
    allend = dt.now()
    print "Total running time: ", allend - allstart
    
    # --- CLONE DELETION PROCESS ---
    # Get contract directory
    contract = Data[contracts]
    contractdir = const.CONTRACTDIR % contract['ProjectID']
    
    # Delete clones directory
    print ">> Deleting clones...",
    existing = os.listdir(contractdir)
    for folder in existing:
        if folder == "clones":
            shutil.rmtree('%s' % contractdir + '\\clones')
    print "Success!"
    return 0

# Excel Report
def run_xl_report(contract):
    '''
    Runs the Excel portion of the final report.
    '''
    # Define global variables
    global temp_path
    global rbd_path
    global rbda_path
    global banners_1
    global banners_2
    global excel_temp
    global word_temp
    global excel_report
    global word_report
    global excel_pdf
    global merge_temp
    global mail_merge
    global prereqs

    # Start Progress bar
    ui.progress_bar()

    print "\nGathering resources...",
    # Get the paths before we start
    temp_path = get_temp_path(contract)
    rbd_path = get_RBD(contract, temp_path)
    rbda_path = get_RBD_A(contract, temp_path)
    banners_1 = get_banners_1(contract)
    banners_2 = get_banners_2(contract)
    excel_temp = get_excel_temp(contract, temp_path)
    word_temp = get_word_temp(contract, temp_path)
    excel_report = get_excel_report(contract)
    word_report = get_word_report(contract)
    excel_pdf = get_excel_pdf(contract)
    merge_temp = get_merge_temp(contract)
    mail_merge = get_mail_merge(contract)
    print "Success!"

    # Verify that the prepwork module was run first.
    print "\nVerifying prerequisites...",
    prereqs = os.path.isdir(temp_path)

    if prereqs is False:
        check_prereqs(prereqs)
    #     print "\n>> You must run the prep work module (cahpsprep.py) before running this module."
    #     print ">> This is because several files are created within the prep work module that are required to successfully create contract reports."
    #     print ">> This program will now quit. Please run the prepwork module first."
    #     print ">> Goodbye!"
    #     sys.exit()
    else:
        print "Success!"
        # Update Verify Check status
        ui.verify_check.config(fg="#60E83B")
        ui.verify_check_text.config(fg="#60E83B")
        time.sleep(0.1)
        ui.progress.step()
        ui.progress.update()

    # Make sure we have all the required inputs
    try:
        assert os.path.exists(rbd_path)
    except AssertionError:
        print "Oops! It looks like you don't have an RBD for this contract."
        print "Try again once you have an RBD available."
    try:
        assert os.path.exists(banners_1)
        assert os.path.exists(banners_2)
    except AssertionError:
        print "Banners not found. Please check the contract directory"
        print "as they may have been moved or renamed."

    # Copy appropriate Excel & Word templates to contract folder
    print "Cloning templates...",
    cp(excel_temp, excel_report)
    cp(word_temp, word_report)
    cp(merge_temp, mail_merge)
    print "Success!"

    #pdb.set_trace()
    # Begin Excel report production
    print "\nBeginning Excel Report production..."
    ui.progress.update()
    with excel.xlconnection() as xlapp:

        # Open the Excel report
        #begin_excel(contract, xlapp)
        print ">> Opening Excel workbook...",
        global xlrpt
        xlrpt = excel.open_workbook(xlapp, excel_report)
        print "Success!"

        # Find the tabs you want to paste data into
        global banners_1_dst
        global banners_2_dst
        global rbd_dst
        global rbd2_dst

        banners_1_dst = excel.get_tab(xlrpt, "Banners1")
        banners_2_dst = excel.get_tab(xlrpt, "Banners2")
        rbd_dst = excel.get_tab(xlrpt, "RBD")
        if contract['SurveyType'] == 'MCS CCC':
            rbd2_dst = excel.get_tab(xlrpt, "RBD2")

        # # Get .csv core banner set 1
        apply_banner1(contract, xlapp)
        # banners_1_src = excel.open_workbook(xlapp, banners_1)
        # banners_1_data_src = excel.get_tab(banners_1_src, "%s Banners1" %
        #                                    contract['ProjectID'])
        # # Copy core banners to Banners1 tab
        # excel.paste_range(banners_1_data_src, "A:Y", banners_1_dst, "A:Y")
        # # Remove '-' from Wincross tab
        # excel.replace_zeros(banners_1_dst, "A:Y")
        # # Close core banner file
        # print ">> Applying Banner Set 1...",
        # del banners_1_data_src
        # excel.close_workbook(banners_1_src, False)
        # print "Success!"

        # # Get .csv core banner set 2
        apply_banner2(contract, xlapp)
        # banners_2_src = excel.open_workbook(xlapp, banners_2)
        # banners_2_data_src = excel.get_tab(banners_2_src, "%s Banners2" %
        #                                    contract['ProjectID'])
        # # Copy core banners to Banners2 tab
        # excel.paste_range(banners_2_data_src, "A:Y", banners_2_dst, "A:Y")
        # # Remove '-' from Wincross tab
        # excel.replace_zeros(banners_2_dst, "A:Y")
        # # Close core banner file
        # print ">> Applying Banner Set 2...",
        # del banners_2_data_src
        # excel.close_workbook(banners_2_src, False)
        # print "Success!"

        # # Get RBD
        apply_rbd(contract, xlapp)
        # if contract['SurveyType'] == 'MCS CCC':
        #     rbdbooksrc = excel.open_workbook(xlapp, rbda_path)
        #     rbdtabsrc = excel.get_tab(rbdbooksrc, "RoundsByDispoCAHPS2017")
        #     rbd2booksrc = excel.open_workbook(xlapp, rbd_path)
        #     rbd2tabsrc = excel.get_tab(rbd2booksrc, "RoundsByDispoCAHPS2017")
        #     excel.paste_range(rbd2tabsrc, "A:AA", rbd2_dst, "A:AA")
        #     del rbd2tabsrc
        #     excel.close_workbook(rbd2booksrc, False)
        # else:
        #     rbdbooksrc = excel.open_workbook(xlapp, rbd_path)
        #     rbdtabsrc = excel.get_tab(rbdbooksrc, "RoundsByDispoCAHPS2017")
        # # Copy RBD to RBD tab
        # excel.paste_range(rbdtabsrc, "A:AA", rbd_dst, "A:AA")
        # # Close RBD file
        # print ">> Applying RBD file...",
        # del rbdtabsrc
        # excel.close_workbook(rbdbooksrc, False)
        # print "Success!"

        # # Group sheets for PDF
        # create_xl_pdf(contract, xlrpt)
        # if contract['SurveyType'] == 'CAS':
        #     allsheets = const.ALL_CAS_SHEETS
        # elif contract['SurveyType'] == 'MAS':
        #     allsheets = const.ALL_MAS_SHEETS
        # elif contract['SurveyType'] == 'MCS':
        #     allsheets = const.ALL_MCS_SHEETS
        # elif contract['SurveyType'] == 'MCS CCC':
        #     allsheets = const.ALL_MCS_CCC_SHEETS
        # else:
        #     return 1
        # print ">> Grouping sheets for PDF report...",
        # group = excel.group_sheets(xlrpt, allsheets)

        # # Export Excel report to PDF
        # print ">> Exporting to PDF...",
        # excel.export_excel_pdf(group, excel_pdf)
        # print "Success!"

        # # Create a MergeData file
        create_merge(contract, xlapp)
        # mergecsv = excel.open_workbook(xlapp, mail_merge)
        # mergedatadst = excel.get_tab(mergecsv, 1)
        # mergedatasrc = excel.get_tab(xlrpt, "MergeData")
        # srcdata = mergedatasrc.Range("1:2")
        # srcdata.Copy()
        # mergedatadst.Range("1:2").PasteSpecial(
        #     Paste=-4122,
        #     Operation=-4142,
        #     SkipBlanks=False,
        #     Transpose=False)
        # mergedatadst.Range("1:2").PasteSpecial(
        #     Paste=-4163,
        #     Operation=-4142,
        #     SkipBlanks=False,
        #     Transpose=False)
        # print ">> Creating MergeData file...",
        # mergecsv.SaveAs(Filename=mail_merge)
        # excel.close_workbook(mergecsv)
        # print "Success!"

        # Cleanup and closing
        print ">> Beginning cleanup procedures...",
        mergedatasrc.Select(Replace=True)
        mergedatasrc.Range("A1").Select()
        excel.close_workbook(xlrpt)

        # Test for correct PDF
        global finrpt
        print "\n>> Opening files to create PDF...",
        finrpt = excel.open_workbook(xlapp, excel_report)
        print "Success!"

        create_xl_pdf(contract, finrpt)

        excel.close_workbook(finrpt)
        excel.quit_xl_app(xlapp)
        print "Success!"



    # End Excel report production
    return 0

def check_module( projectid):
    '''
    Check process status and update labels.
    '''
    ui.module_check.config(fg="#B68C14")
    ui.module_check_text.config(fg="#B68C14")
    ui.update_idletasks()
    
    pass_id = projectid
    status = start_reports(pass_id)

    time.sleep(0.1)
    ui.progress.update()

    if status == 0:
        # Update Run Module status
        ui.module_check.config(fg="#60E83B")
        ui.module_check_text.config(fg="#60E83B")
        ui.update_idletasks()
    else:
        # Update Run Module status
        ui.module_check.config(fg="#E8483B")
        ui.module_check_text.config(fg="#E8483B")
        ui.mod_passfail.set(ui.failpass)
        ui.update_idletasks()

def begin_excel(contract):
    '''
    This function updates the UI through the functions.
    '''
    # Update Excel Start Check status
    ui.ex_start_check.config(fg="#B68C14")
    ui.ex_start_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    contract = contract
    
    status = run_xl_report(contract)

    if status == 0:
        # Update Excel Start Check status
        ui.ex_start_check.config(fg="#60E83B")
        ui.ex_start_check_text.config(fg="#60E83B")
        ui.update_idletasks()
    else:
        # Update Excel Start Check status
        ui.ex_start_check.config(fg="#E8483B")
        ui.ex_start_check_text.config(fg="#E8483B")
        ui.ex_start_passfail.set(ui.failpass)
        ui.update_idletasks()

def apply_banner1(contract, xlapp):
    '''
    This function updates the UI through the functions.
    '''
    # Update Banner 1 Check status
    ui.banner1_check.config(fg="#B68C14")
    ui.banner1_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    

    try:
        # Get .csv core banner set 1
        banners_1_src = excel.open_workbook(xlapp, banners_1)
        banners_1_data_src = excel.get_tab(banners_1_src, "%s Banners1" %
                                        contract['ProjectID'])

        # Copy core banners to Banners1 tab
        excel.paste_range(banners_1_data_src, "A:Y", banners_1_dst, "A:Y")
        # Remove '-' from Wincross tab
        excel.replace_zeros(banners_1_dst, "A:Y")
        # Close core banner file
        print ">> Applying Banner Set 1...",
        del banners_1_data_src
        excel.close_workbook(banners_1_src, False)
        print "Success!"
    except IOError:
        # Update Banner 1 Check status
        ui.banner1_check.config(fg="#E8483B")
        ui.banner1_check_text.config(fg="#E8483B")
        ui.banner1_passfail.set(ui.failpass)
        ui.update_idletasks()
    
    # Update Banner 1 Check status
    ui.banner1_check.config(fg="#60E83B")
    ui.banner1_check_text.config(fg="#60E83B")
    ui.update_idletasks()

def apply_banner2(contract, xlapp):
    '''
    This function updates the UI through the functions.
    '''
    # Update Banner 2 Check status
    ui.banner2_check.config(fg="#B68C14")
    ui.banner2_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    
    try:
        # Get .csv core banner set 2
        banners_2_src = excel.open_workbook(xlapp, banners_2)
        banners_2_data_src = excel.get_tab(banners_2_src, "%s Banners2" %
                                        contract['ProjectID'])
        # Copy core banners to Banners2 tab
        excel.paste_range(banners_2_data_src, "A:Y", banners_2_dst, "A:Y")
        # Remove '-' from Wincross tab
        excel.replace_zeros(banners_2_dst, "A:Y")
        # Close core banner file
        print ">> Applying Banner Set 2...",
        del banners_2_data_src
        excel.close_workbook(banners_2_src, False)
        print "Success!"
    except:
        # Update Banner 2 Check status
        ui.banner2_check.config(fg="#E8483B")
        ui.banner2_check_text.config(fg="#E8483B")
        ui.banner2_passfail.set(ui.failpass)
        ui.update_idletasks()
    
    # Update Banner 2 Check status
    ui.banner2_check.config(fg="#60E83B")
    ui.banner2_check_text.config(fg="#60E83B")
    ui.update_idletasks()

def apply_rbd(contract, xlapp):
    '''
    This function updates the UI through the functions.
    '''
    # Update RBD Check status
    ui.rbd_check.config(fg="#B68C14")
    ui.rbd_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    
    try:
        # Get RBD
        if contract['SurveyType'] == 'MCS CCC':
            rbdbooksrc = excel.open_workbook(xlapp, rbda_path)
            rbdtabsrc = excel.get_tab(rbdbooksrc, "RoundsByDispoCAHPS2016")
            rbd2booksrc = excel.open_workbook(xlapp, rbd_path)
            rbd2tabsrc = excel.get_tab(rbd2booksrc, "RoundsByDispoCAHPS2016")
            excel.paste_range(rbd2tabsrc, "A:AE", rbd2_dst, "A:AE")
            del rbd2tabsrc
            excel.close_workbook(rbd2booksrc, False)
        else:
            rbdbooksrc = excel.open_workbook(xlapp, rbd_path)
            rbdtabsrc = excel.get_tab(rbdbooksrc, "RoundsByDispoCAHPS2016")
        # Copy RBD to RBD tab
        excel.paste_range(rbdtabsrc, "A:AE", rbd_dst, "A:AE")
        # Close RBD file
        print ">> Applying RBD file...",
        del rbdtabsrc
        excel.close_workbook(rbdbooksrc, False)
        print "Success!"
    except IOError:
        # Update RBD Check status
        ui.rbd_check.config(fg="#E8483B")
        ui.rbd_check_text.config(fg="#E8483B")
        ui.rbd_passfail.set(ui.failpass)
        ui.update_idletasks()
    
    # Update RBD Check status
    ui.rbd_check.config(fg="#60E83B")
    ui.rbd_check_text.config(fg="#60E83B")
    ui.update_idletasks()

def create_xl_pdf(contract, xlrpt):
    '''
    This function updates the UI through the functions.
    '''
    # Update Creat Excel PDF Check status
    ui.ex_pdf_check.config(fg="#B68C14")
    ui.ex_pdf_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    
    # Group sheets for PDF
    if contract['SurveyType'] == 'CAS':
        allsheets = const.ALL_CAS_SHEETS
    elif contract['SurveyType'] == 'MAS':
        allsheets = const.ALL_MAS_SHEETS
    elif contract['SurveyType'] == 'MCS':
        allsheets = const.ALL_MCS_SHEETS
    elif contract['SurveyType'] == 'MCS CCC':
        allsheets = const.ALL_MCS_CCC_SHEETS
    else:
        return 1
        # Update RBD Check status
        ui.ex_pdf_check.config(fg="#E8483B")
        ui.ex_pdf_check_text.config(fg="#E8483B")
        ui.ex_pdf_passfail.set(failpass)
        ui.update_idletasks()
    print ">> Grouping sheets for PDF report...",
    group = excel.group_sheets(xlrpt, allsheets)

    # Export Excel report to PDF
    print ">> Exporting to PDF...",
    excel.export_excel_pdf(group, excel_pdf)
    print "Success!"
    
    # Update RBD Check status
    ui.ex_pdf_check.config(fg="#60E83B")
    ui.ex_pdf_check_text.config(fg="#60E83B")
    ui.update_idletasks()

def create_merge(contract, xlapp):
    '''
    This function updates the UI through the functions.
    '''
    print mail_merge

    # Update MergeData Check status
    ui.merge_create_check.config(fg="#B68C14")
    ui.merge_create_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    
    try:
        # Create a MergeData file
        global mergecsv
        global mergedatadst
        global mergedatasrc
        global srcdata

        mergecsv = excel.open_workbook(xlapp, mail_merge)
        mergedatadst = excel.get_tab(mergecsv, 1)
        mergedatasrc = excel.get_tab(xlrpt, "MergeData")
        srcdata = mergedatasrc.Range("1:2")
        srcdata.Copy()
        mergedatadst.Range("1:2").PasteSpecial(
            Paste=-4122,
            Operation=-4142,
            SkipBlanks=False,
            Transpose=False)
        mergedatadst.Range("1:2").PasteSpecial(
            Paste=-4163,
            Operation=-4142,
            SkipBlanks=False,
            Transpose=False)
        print ">> Creating MergeData file...",
        mergecsv.SaveAs(Filename=mail_merge)
        excel.close_workbook(mergecsv)
        print "Success!"
    except IOError:
        # Update MergeData Check status
        ui.merge_create_check.config(fg="#E8483B")
        ui.merge_create_check_text.config(fg="#E8483B")
        ui.merge_create_passfail.set(ui.failpass)
        ui.update_idletasks()
    
    # Update MergeData Check status
    ui.merge_create_check.config(fg="#60E83B")
    ui.merge_create_check_text.config(fg="#60E83B")
    ui.update_idletasks()

def begin_word(contract):
    '''
    This function updates the UI through the functions.
    '''
    # Update Word Start Check status
    ui.wrd_start_check.config(fg="#B68C14")
    ui.wrd_start_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()

    contract = contract
    
    status = run_wrd_report(contract)

    if status == 0:
        # Update Word Start Check status
        ui.wrd_start_check.config(fg="#60E83B")
        ui.wrd_start_check_text.config(fg="#60E83B")
        ui.update_idletasks()
    else:
        # Update Word Start Check status
        ui.wrd_start_check.config(fg="#E8483B")
        ui.wrd_start_check_text.config(fg="#E8483B")
        ui.wrd_start_passfail.set(ui.failpass)
        ui.update_idletasks()

def run_merge(wrdapp):
    '''
    This function updates the UI through the functions.
    '''
    # Update Run Merge Check status
    ui.run_merge_check.config(fg="#B68C14")
    ui.run_merge_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    
    try:
        # Open the Word report
        print "\n>> Opening Word document...",
        global wrdrpt
        wrdrpt = word.open_word_doc(wrdapp, word_report)
        print "Success!"

        # Run Mail Merge
        print ">> Run Mail Merge...",
        word.mail_merge(wrdrpt, mail_merge)
        print "Success!"
    except IOError:
        # Update Run Merge Check status
        ui.run_merge_check.config(fg="#E8483B")
        ui.run_merge_check_text.config(fg="#E8483B")
        ui.run_merge_passfail.set(ui.failpass)
        update_idletasks()
    
    # Update Run Mail Merge Check status
    ui.run_merge_check.config(fg="#60E83B")
    ui.run_merge_check_text.config(fg="#60E83B")
    ui.update_idletasks()


def create_wrd_pdf(wrdapp):
    '''
    This function updates the UI through the functions.
    '''
    # Update Create Word PDF Check status
    ui.wrd_pdf_check.config(fg="#B68C14")
    ui.wrd_pdf_check_text.config(fg="#B68C14")
    time.sleep(0.1)
    ui.progress.step()
    ui.progress.update()
    
    try:
        # Update Table of Contents
        print ">> Updating Table of Contents...",
        word.update_toc(wrdrpt)
        print "Success!"

        # Export Word report to PDF
        print ">> Exporting to PDF...",
        word.export_word_pdf(wrdrpt, word_pdf)
        print "Success!"
    except IOError:
        # Update Create Word PDF Check status
        ui.wrd_pdf_check.config(fg="#E8483B")
        ui.wrd_pdf_check_text.config(fg="#E8483B")
        ui.wrd_pdf_passfail.set(ui.failpass)
        ui.update_idletasks()
    
    # Update Create Word PDF Check status
    ui.wrd_pdf_check.config(fg="#60E83B")
    ui.wrd_pdf_check_text.config(fg="#60E83B")
    ui.update_idletasks()

# Word Report
def run_wrd_report(contract):
    '''
    Runs the Word portion of the final report.
    '''
    # Define global variables
    global word_report
    global word_pdf
    global mail_merge
    
    # Get the paths before we start
    print "Gathering resources...",
    word_report = get_word_report(contract)
    word_pdf = get_word_pdf(contract)
    mail_merge = get_mail_merge(contract)
    print "Success!"

    # Begin Word report production
    print "Begin Word Report production..."
    ui.progress.update()
    with word.wrdconnection() as wrdapp:

        # # Open the Word report
        run_merge(wrdapp)
        # print "\n>> Opening Word document...",
        # wrdrpt = word.open_word_doc(wrdapp, word_report)
        # print "Success!"

        # # Run Mail Merge
        # print ">> Run Mail Merge...",
        # word.mail_merge(wrdrpt, mail_merge)
        # print "Success!"

        # # Update Table of Contents
        create_wrd_pdf(wrdapp)
        # print ">> Updating Table of Contents...",
        # word.update_toc(wrdrpt)
        # print "Success!"

        # # Export Word report to PDF
        # print ">> Exporting to PDF...",
        # word.export_word_pdf(wrdrpt, word_pdf)
        # print "Success!"

        # Cleanup and closing
        print ">> Beginning cleanup procedures...",
        word.close_word_doc(wrdrpt)
        word.quit_word_app(wrdapp)
        print "Success!"

    # End Word report production
    return 0

# @utils.timer_dec


def merge_components(contract):
    '''
    Merges the Excel and Word PDFs together depending on the report type.
    '''

    # Get paths before we start.
    cq_book_pdf = get_cq_book_pdf(contract)
    final_pdf = get_final_pdf(contract)
    excel_pdf = get_excel_pdf(contract)
    word_pdf = get_word_pdf(contract)

    if os.path.exists(cq_book_pdf):
        cqbookpath = cq_book_pdf
    else:
        cqbookpath = None

    if contract['SurveyType'] == 'CAS':
        merge_report(final_pdf,
                    word_pdf,
                    excel_pdf,
                    pdfmerge.CAS_WORD_SECTIONS,
                    pdfmerge.CAS_EXCEL_SECTIONS,
                    cqbookpath=cqbookpath)
    elif contract['SurveyType'] == 'MAS':
        merge_report(final_pdf,
                    word_pdf,
                    excel_pdf,
                    pdfmerge.MAS_WORD_SECTIONS,
                    pdfmerge.MAS_EXCEL_SECTIONS,
                    cqbookpath=cqbookpath)
    elif contract['SurveyType'] == 'MCS':
        merge_report(final_pdf,
                    word_pdf,
                    excel_pdf,
                    pdfmerge.MCS_WORD_SECTIONS,
                    pdfmerge.MCS_EXCEL_SECTIONS,
                    cqbookpath=cqbookpath)
    elif contract['SurveyType'] == 'MCS CCC':
        merge_report(final_pdf,
                    word_pdf,
                    excel_pdf,
                    pdfmerge.MCS_CCC_WORD_SECTIONS,
                    pdfmerge.MCS_CCC_EXCEL_SECTIONS,
                    cqbookpath=cqbookpath)

    else:
        return 1
    return 0

def check_prereqs(prereqs):
    '''
    Verify that the clones directory exists.
    '''
    no_prereq = None
    if prereqs is False:
        ui.output_text.insert(tk.INSERT, "\n>> You must run the prep work module (cahpsprep.py) before running this module." + \
                                            "\n>> This is because several files are created within the prep work module that " + \
                                            "\nare required to successfully create contract reports.")
        no_prereq = tkMessageBox.showerror("No prerequisites!", "Sorry, you must run the prep work module prior to running this module.")
        print no_prereq

        if no_prereq == 'ok':
            sys.exit()
        

            
            
    



