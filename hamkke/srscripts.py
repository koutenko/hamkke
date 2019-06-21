#! /usr/bin/env python
"""
Module Name
===========
 $ srscripts.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Shell Syntax
============
 N/A

Description
===========
Here we define the SRScript class.  There is only one method so far, and that
is to run a .mrs file.  The run method should be abstract enough to be useful
in many different scenarios, but not so abstract that it confuses the user.

"""

import sys
import os
from shutil import copyfile
import shlex
import Tkinter as tk
from subprocess import Popen, PIPE
import hamkke.errors as errors

class SRScript(object):
    '''
    Defines the SRScript object class.
    '''

    def __init__(self, mrsfile, stype=None, constants=None, script=None, bannerpath=None,
                 mddfile=None, dmstemplate=None):
        '''
        Initialize the SRScript object.

        mrsfile : String filepath of the .mrs file to run. (Required)

        stype : Optional argument to identify the script type (i.e. 'core',
                or 'custom' or 'additional segmentation')

        constants : An optional dictionary of string constants to give the
                    mrScriptCL program as arguments. For example, a CAHPS
                    constants dictionary might look like this:
                        {'ProjectID' : '1010101',
                         'PlanName' : 'Example Health Plan',
                         'SurveyType' : 'CAS'
                         'Q1Name' : 'Example Health Plan'}
                    It is the responsibility of the user to ensure that the
                    set of constants given coincide with those required by
                    the .mrs script file.

        bannerpath : Optional string filepath to location where resulting
                     banners will live.  This is used for copying the TMG
                     logo to the right place.
        
        text :       Connection to the output window in the GUI.
        '''

        # Assign given arguments to variables.
        self.mrsfile = mrsfile
        self.stype = stype
        self.constants = constants
        self.logopath = r"\\10.10.210.24\PM2018_Peak10\Analytics CAHPS" + \
                        r"\Report Templates\Banners\logo.png"
        self.script = script
        self.bannerpath = bannerpath
        #self.text = text

    def run(self, debug=False):
        '''
        Run any Survey Reporter script.
        Errors aren't handled particularly well, but it will output an error
        message if something goes wrong.

        debug (bool) : Set to True to print some additional information.
        '''
        # Make sure the mrsfile argument is actually a file
        assert os.path.isfile(self.mrsfile)

        mrsd, mrsf = os.path.split(os.path.abspath(self.mrsfile))
        currdir = os.getcwd()
        os.chdir(mrsd)

        # Set the required mrScriptCL arguments and add constants as needed
        args = '%s' % self.script + ' \"%s\"' % self.mrsfile

        if self.constants:
            for key, item in self.constants.iteritems():
                args += ' /a:%s=\"%s\"' % (key, item)
        if not debug:
            args += " /s"

        if debug:
            print ">>>> TEMPLATE ARGUMENTS TEST <<<<"
            print "\nArguments: "
            print args, '\n'

        # Tell the user what's going on
        if self.stype:
            print ">> Running %s script..." % self.stype,
        else:
            print ">> Running script...",
        # Run the mrScriptCL program as a subprocess
        try:
            run = Popen(args, stdout=PIPE, stderr=PIPE)
            (out, err) = run.communicate()
            run.wait()
            out = str(out).strip("\t\n\r")
            err = str(err).strip("\t\n\r")
            if debug:
                print '\n', out
                print err
            else:
                if out:
                    print errors.ERROR_CODE_MAP[1]
                    raise errors.ScriptError(out)
                if err:
                    print errors.ERROR_CODE_MAP[1]
                    raise errors.ScriptError(err)
                else:
                    print errors.ERROR_CODE_MAP[0]
        except errors.ScriptError as scripterr:
            print scripterr
        finally:
            os.chdir(currdir)

        # If the projectpath attribute exists, copy the TMG logo to it
        if self.bannerpath:
            copyfile(self.logopath, self.bannerpath + r"\logo.png")

        return 0
