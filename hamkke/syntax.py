#! /usr/bin/env python
"""
Module Name
===========
 $ syntax.py

Author
======
 Jimin McClain
 Jr. Data Analyst

Shell Syntax
============
 N/A

Description
===========
Here we define the Syntax object.  At this point, we only have a run method
after the initialization (as well as a run_test method), but stay tuned for
more updates here.

The Syntax object collects the arguments passed through from the prep_work method
and creates an SPJ file which will then be used to run the syntax via the subprocess.
"""

import os
import shlex
import hamkke as ra
from hamkke.cahps import constants as const
from hamkke.mcahps import constants as mconst
from subprocess import Popen, PIPE
import hamkke.errors as errors
from jinja2 import Environment, PackageLoader
import Tkinter as tk
import pdb
import time

class Syntax(object):
    '''
    This is the base Syntax class.  Most of the methods here will be overloaded
    by child classes, but this should serve as a basic framework.

    spj (str) : Optional path to SPSS Production Facility .spj file.  If the
                spj argument is omitted, a default is used which will run the
                syntax file denoted by the sps argument.
    sps (str) : Optional path to SPSS syntax file.  If no valid spj file is
                provided, this is a required argument.
    spv (str) : Optional path to SPSS output file.
    stype (str) : Optional syntax type (e.g. 'core' or 'custom')
    symbols (dict) : Dictionary of placeholders (e.g. {'@ProjectID', '1010101'})

    '''

    # Initialise Syntax object.
    def __init__(self, projectid=None, spj=None, sps=None, spv=None, stype=None, symbols=None, stats=None, debug=False):
        if debug:
            print "Project ID: ", projectid
            print "SPJ: ", spj
            print "SPS: ", sps
            print "SPV: ", spv
            print "SType: ", stype
            print "Symbols: ", symbols
            print "Stats Path", stats
            #print "Output text", text

        if projectid is not None:
            self.projectid = projectid

        if stats is not None:
            self.stats = os.path.abspath(stats)

        if sps is not None:
            self.sps = os.path.abspath(sps)
        else:
            self.sps = sps

        if spv is not None:
            self.spv = os.path.abspath(spv)
        else:
            self.spv = spv

        self.stype = stype
        self.symbols = symbols

        if os.path.exists(os.path.abspath(spj)):
            self.spj = os.path.abspath(spj)
        else:
            self.spj = self.create_spj(spj)

#        # Make some assertion so we don't get too far with bad data
#        if self.spj is None:
#            assert os.path.isfile(self.sps)
#        else:
#            # Let's only check that self.spj *could* be a valid file (even if
#            # it does not exist yet).  This will allow us to use the spj
#            # attribute as the default outfile argument for the create_spj
#            # method.  Do to that, we'll check that the directory exists and
#            # that the file extension is .spj.
#            assert self.spj[-4:].lower() == '.spj'

        if not os.path.isfile(self.spj):
            self.spj = self.create_spj()
            return True

        if self.spv:
            # Here we do the same kind of check we did with self.spj.  It
            # doesn't need to exist yet, but it should be a valid possibility.
            spv_dir = os.path.dirname(self.spv)
            assert os.path.isdir(spv_dir)
            assert self.spv[-4:].lower() == '.spv'

        if self.symbols:
            assert type(self.symbols) is dict
            for symbol in self.symbols:
                # Symbols have to be prefixed with '@'.
                assert symbol[0] == '@'

        if self.stype:
            assert type(self.stype) is str

    def create_spj(self, outfile=None, debug=False):
        '''
        Use the jinja2 templating framework to make an SPSS SPJ file based
        on a default template.  This requires that a valid SPS syntax file has
        been defined as the self.sps attribute.

        outfile (str) : Optional file path to location where the created SPJ
                        file should go.  Defaults to self.spj if provided.  If
                        self.spj is not defined, then the outfile argument is
                        required.
        '''
        # Make sure we have the required information before we start
        assert self.sps is not None
        if outfile is None and hasattr(self, 'spj'):
            outfile = self.spj
        if outfile is None and not hasattr(self, 'spj'):
            outfile = os.path.join(os.path.dirname(self.sps), 'tmp.spj')

        # assert outfile[-4:].lower() == '.spj'
        assert os.path.isdir(os.path.dirname(os.path.abspath(outfile)))

        print ">> Creating SPJ file...",
        env = Environment(loader=PackageLoader('hamkke',
                                               'templates'))
        template = env.get_template('default.spj')
        with open(outfile, 'w') as out:
            out.write(template.render(spv=self.spv, sps=self.sps,
                                      symbols=self.symbols))
        print "Success!"
        return outfile

    def get_args(self, debug=False):
        '''
        Get the arguments all good and gathered before we move on.
        '''
        print ">> Gathering resources...",
        stats = self.stats
        spj_dir, spj_file = os.path.split(self.spj)
        currdir = os.getcwd()
        os.chdir(spj_dir)
        print "Success!\n"

        args = "%s" % stats +  " %s" % spj_file + " -production silent"
        if self.symbols:
            args += " -symbol"

            for key, val in self.symbols.iteritems():
                args += " %s" % key
                args += ' "%s"' % val

        if debug:
            print '\n', args

        return args
    
    def check_syntax(self, debug=False):
        '''
        Checks the Syntax folder if there is a Java error to see if the process worked.

        Ideally a method that can be phased out if the source of the Java error can be discovered.
        '''
        projectid = self.projectid
        symbols = self.symbols
        mod_type = None
        test = None

        if '@ProjectID' in symbols:
            mod_type = "CAHPS"
            syn_path = const.SYNTAX_PATH % projectid
            data_path = const.DATA_PATH % projectid
        else:
            mod_type = "MCAHPS"
            syn_path = mconst.SYNTAX_PATH % projectid
            data_path = mconst.DATA_PATH % projectid

        # CAHPS Output Test
        spv_test = "\\%s" % projectid + ".spv"
        pdf_test = "\\%s" % projectid + "_Core_Output.pdf"
        core_spv_test = "\\%s" % projectid + "_Core_Output.spv"
        job_test = "\\%s" % projectid + "syntax.spj"
        final_test = "\\%s" % projectid + "_FinalData.sav"

        # MCAHPS Output Test
        mcahps_core_spv = "\\%s" % projectid + " Core Output.spv"
        mcahps_core_pdf = "\\%s" % projectid + " Core Output.pdf"
        mcahps_recoded = "\\%s" % projectid + "_RecodedData.sav"

        time.sleep(3)

        if test is None:
            if mod_type is "CAHPS":
                test = os.path.isfile(syn_path + spv_test)
                test = os.path.isfile(syn_path + pdf_test)
                test = os.path.isfile(syn_path + core_spv_test)
                test = os.path.isfile(syn_path + job_test)
                test = os.path.isfile(data_path + final_test)
            else:
                test = os.path.isfile(syn_path + mcahps_core_spv)
                test = os.path.isfile(syn_path + mcahps_core_pdf)
                test = os.path.isfile(data_path + mcahps_recoded)
        
        return test

    def run(self, debug=False):
        '''
        Run the spj and/or sps file designated in the object initialization.

        The subprocess calls the stats.exe program which runs the file using SPSS.
        Here the stats variable contains the path to the local directory containing the
        stats.exe program on the LOCAL MACHINE.
        '''
        args = self.get_args()
        currdir = os.getcwd()

        if self.stype:
            print "Running %s syntax..." % self.stype,
        else:
            print "Running syntax...",
        
        run = Popen(args, shell=False, stdout=PIPE, stderr=PIPE)
        (out, err) = run.communicate()
        run.wait()
        out = str(out).strip("\t\n\r")
        err = str(err).strip("\t\n\r")
        java_err = False
        syn_test = None
        if out:
            print out
        if err:
            print "\n>> Funny noises? Something to look into. But let's see how this goes, shall we?"
            java_err = True
            syn_test = self.check_syntax()
        if debug:
            print err
        if java_err is not True:
            print errors.ERROR_CODE_MAP[0]

        if syn_test is True:
            print ">> All syntax and final data files created successfully. No worries!"
            return 0
        else:
            print ">> One or more syntax or data files may not have been created." 
            print ">> Please check the project directory before proceeding."
            return 1

        os.chdir(currdir)
