#!/usr/bin/env python
"""
Module Name
===========
 $ datamgmt.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Shell Syntax
============
 N/A

Description
===========
The datamgmt module is meant for high-level data maintenance.  This includes
setting up a desired reporting directory structure, archiving old data, and
moving files from place to place.

"""

import os
import sys
import shutil
import zipfile
import hamkke as ra
import Tkinter as tk
import tkMessageBox
import time
from datetime import datetime as dt

# ARCHIVE FILES CHOICE DIALOG
def skip_choice(gui_toggle):
    '''
    Function for the verification dialog that occurs
    to ask the user if they want to archive existing files or not.
    This function is for the GUI modules.
    '''
    arch_skip = None

    if gui_toggle is True:
        arch_skip = tkMessageBox.askyesnocancel("Archive files..?", "The Reporting directory for this project already exists.\n" + \
                                                "Do you want to archive this data?\n\n" + \
                                                "WARNING: Typing 'no' will overwrite existing data. Choose wisely!")
    else:
        print "The Reporting directory for this project already exists. Do you want to archive this data?\n"

        skip = str(raw_input("WARNING: Typing 'no' will overwrite existing data. Choose wisely or type 'quit' to exit!: "))
        if skip.lower() not in ['yes', 'y', 'no', 'n']:
            print "\n>>> Please type 'yes' or 'no'."
        if skip.lower() in ['yes', 'y']:
            arch_skip == True
        if skip.lower() in ['no', 'n']:
            arch_skip == False
        if skip.lower() in ['quit', 'q']:
            arch_skip == None

    if arch_skip is True:
        return True
    if arch_skip is None:
        print "\nFascinating. Lovely chat.\n"
        print "Off you go!\n"
        sys.exit()
    else:
        return False

def makedirectory(basedir, projectID, gui_toggle, rptfolders=["Archive", "Banners", "Data", "Syntax"]):
    '''
    Makes the report's reporting directory, if one does not already exist.
    The reporting directory is laid out as follows:

        basedir/
            [project]/
                Archive/
                Banners/
                Data/
                Final Report/
                Syntax/

    Basedir is a string representing the filepath to the directory in which
    a project's reporting should be held.  For example, the 2014 MCAHPS
    basedir would be r"\\mercury\pm2014\Analytics Medicare CAHPS\Reporting".
    ProjectID is the unique identifier used to label the project.  In most
    cases, this would be the SurveyID.  For MCAHPS, this is the contract
    number.  The optional rptfolders argument is a list of folders to create
    in the project's reporting folder.  In order to use the archivefiles
    method, this list `must` include an Archive folder.
    '''
    assert os.path.isdir(basedir)
    assert isinstance(rptfolders, list)   
    projectID = str(projectID)
    choice = None
    gui_toggle = gui_toggle

    path = os.path.join(basedir, projectID)

    # If the reporting path does not already exist, make it

    if not os.path.exists(path):
        os.makedirs(path)
    
        existing = os.listdir(path)
        for folder in rptfolders:
            if folder not in existing:
                os.mkdir(os.path.join(path, folder))

    # If it does, prompt the user to see if they want to archive the files
    else:
        choice = skip_choice(gui_toggle)

    if choice is True:
        existing = os.listdir(path)
        for folder in rptfolders:
            if folder not in existing:
                os.mkdir(os.path.join(path, folder))
        archivefiles(path)

        # Now that you've archived and deleted the old folders, make new ones.
        existing = os.listdir(path)
        for folder in rptfolders:
            if folder not in existing:
                os.mkdir(os.path.join(path, folder))
            else:
                print ">> Archiving old files... Success!"
                # Delays program execution for 3 seconds to ensure the folders have time to populate before continuing
                time.sleep(3)

    # If the user chooses 'no', do not archive the files
    if choice is False:
        print "\n>> Now activating YOLO Mode..."
        existing = os.listdir(path)
        for folder in existing:
            if folder != "Archive":
                shutil.rmtree('%s' % path + '\\' + '%s' % folder)

        # Now that you've deleted the old folders, make new ones.
        existing = os.listdir(path)
        for folder in rptfolders:
            if folder not in existing:
                os.mkdir(os.path.join(path, folder))
            else:
                print ">> Data archiving process skipped... Success!"
                # Delays program execution for 3 seconds to ensure the folders have time to populate before continuing
                time.sleep(3)

    return 0


def archivefiles(path):
    '''
    Archives the current contents of the project reporting folder into a zip
    file labeled with the date and time.

    basedir (str): Path to project reporting folder.

    At this point, there must exist and "Archive" folder in the base directory
    in order for this to work.  If one does not exists, an AssertionError is
    raised.
    '''
    assert os.path.isdir(path)
    rptfolders = os.listdir(path)
    assert isinstance(rptfolders, list)
    assert "Archive" in rptfolders
    rptfolders.remove("Archive")
    currdir = os.getcwd()

    if "tmp" in rptfolders:
        os.chdir(path)
        shutil.rmtree("tmp")
        os.chdir(currdir)

    def skiparchive(path, names):
        return [os.path.join(path, "Archive")]

    print ">> Archiving old files...",

    os.chdir(path)

    # Copy directory structure into a temporary folder.
    for folder in rptfolders:
        shutil.copytree('%s' % folder, 'tmp/%s' % folder)
    print ">> Starting copypasta..."

    # Keep the old data in a zip file labeled with the current date/time.
    zipname = dt.strftime(dt.today(), "%m_%d_%y__%I_%M%p") + ".zip"
    zipf = zipfile.ZipFile(zipname, "w")
    zipdir("tmp", zipf)
    zipf.close()
    shutil.move(zipname, r"Archive\%s" % zipname)
    print ">> Copypasta... Complete!"

    assert isinstance(rptfolders, list)
    shutil.rmtree('tmp')
    try:
        rptfolders.remove('CQ')
    except ValueError:
        pass
    for folder in rptfolders:
        if folder != "Archive":
            shutil.rmtree('%s' % folder)

    os.chdir(currdir)
    return 0


def zipdir(path, zipf):
    '''
    Helper function for zipping up the files in a directory.

    path (str) : file path to zip up
    zipf (zipfile.ZipFile) : writable ZipFile instance
    '''
    for root, _, files in os.walk(path):
        zipf.write(root)
        for filename in files:
            zipf.write(os.path.join(root, filename))


def copytree(src, dst, symlinks=False, ignore=None):
    '''
    Helper function to copy a directory.  Similar to shutil.copytree.
    Thanks to SO user atzz for this.

    src (str) : path to directory to copy from
    dst (str) : path to directory to copy to
    symlinks (bool) : follow symbolic links?
    ignore (func) : function that yields a list of subdirectories to ignore
    '''
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)


def cp(src, dst):
    '''
    Just renamed the shutil.copy2 function to give it a Unix-like name.

    src (str) : path to file to copy
    dst (str) : path to destination for src
    '''
    shutil.copy2(src, dst)

def copy_sps_template(spstemp, temp_path, contractdir):
    '''
    Copy the .sps template file into a UNIQUE temporary folder on the server.
    '''
    sps_file = spstemp
    temp_path = temp_path
    contractdir = contractdir

    # Tests to make sure the paths are correct and the files exist.
    assert os.path.exists(sps_file)

    # Get path to current working directory
    currdir = os.getcwd()

    # Split paths
    source_dir, file_name = os.path.split(sps_file)

    # Move to contract folder
    os.chdir(contractdir)  

    # Create temporary directory to hold the files if it doesn't exist already
    if not os.path.exists(temp_path):
        os.makedirs(temp_path)

    # Make sure temporary directory successfully exists
    assert os.path.isdir(temp_path)

    # Copy file
    cp(sps_file, temp_path)

    # Construct final path to temporary file
    sps_copy = os.path.join(temp_path, file_name)

    print ">> Cloned .sps template file... Success!"
    os.chdir(currdir)
    return sps_copy

def copy_mrs_template(mrstemp, temp_path, contractdir):
    '''
    Copy the .sps template file into a UNIQUE temporary folder on the server.
    '''
    mrs_file = mrstemp
    temp_path = temp_path
    contractdir = contractdir

    # Tests to make sure the paths are correct and the files exist.
    assert os.path.exists(mrs_file)

    # Get path to current working directory
    currdir = os.getcwd()

    # Split paths
    source_dir, file_name = os.path.split(mrs_file)

    # Move to contract folder
    os.chdir(contractdir)  

    # Create temporary directory to hold the files if it doesn't exist already
    if not os.path.exists(temp_path):
        os.makedirs(temp_path)

    # Make sure temporary directory successfully exists
    assert os.path.isdir(temp_path)

    # Copy file
    cp(mrs_file, temp_path)

    # Construct final path to temporary file
    mrs_copy = os.path.join(temp_path, file_name)

    print "\n>> Cloned .mrs template file... Success!"

    os.chdir(currdir)
    return mrs_copy

def aggdirectory(aggdir, projectID, rptfolders=["Files", "Templates"]):
    '''
    Makes a directory for the Aggregate file so that copies of the text data
    can be stored in the Archive folder.
    '''
    #assert os.path.isdir(basedir)
    #assert isinstance(rptfolders, list)
    projectID = str(projectID)

    path = os.path.join(aggdir, projectID)
    #arch_path = path + "\\Archive"

    # If the reporting path does not already exist, make it

    if not os.path.exists(path):
        os.makedirs(path)
    
        existing = os.listdir(path)
        for folder in rptfolders:
            if folder not in existing:
                os.mkdir(os.path.join(path, folder))
                print "Directory created at:"
                print path
            else:
                print "Directory already exists, moving on..."

def rerunbanners(basedir, projectID, rptfolder=["Banners"]):
    '''
    Deletes and recreates the Banners folder so the Banners can be rerun
    without having to redo all of the prepwork process.
    '''
    projectID = str(projectID)
    path = os.path.join(basedir, projectID)

    # First, delete the existing Banners folder and its contents
    print "\n>> Deleting old Banners..."
    existing = os.listdir(path)
    for folder in existing:
        if folder == "Banners":
            shutil.rmtree('%s' % path + '\\Banners')

        # Now that you've deleted the old folder, make a new one.
        existing = os.listdir(path)
        for folder in rptfolders:
            if folder not in existing:
                os.mkdir(os.path.join(path, folder))
            else:
                print ">> Old Banners removal... Success!"
                # Delays program execution for 3 seconds to ensure the folders have time to populate before continuing
                time.sleep(3)
