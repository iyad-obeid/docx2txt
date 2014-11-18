#!/usr/bin/python

# DOCX2TXT.PY
#
#   Iyad Obeid, 11/18/2014, v1.0.2
#
#   Converts docx to text
#   Run with -h or -help flag for more information on how to run
#
#   Code is based on docx.py which is downloaded from here:
#   https://github.com/mikemaccana/python-docx
#
#   Installation requires 
#     apt-get install libxml2-dev libxslt1-dev python-dev
#       (this may be required on linux, shouldn't be necessary on later
#        model osx systems)
#
#     sudo pip install lxml
#     sudo pip install Pillow (formerly PIL)

import sys
import string
import os.path
from docx import opendocx, getdocumenttext

def main():

    # initialize flags
    headerFlag = True
    footerFlag = True
    bodyFlag   = True
    verboseFlag = False
    helpFlag = False

    nArguments = len(sys.argv)-1

    # check all the input switches in order to set up process flow properly
    for i in range( 1 , len(sys.argv) ):

        if (sys.argv[i].lower() == '-noheader') or \
                (sys.argv[i].lower() == '-nohdr') :
            headerFlag = False
            nArguments -= 1

        elif (sys.argv[i].lower() == '-nofooter') or \
                (sys.argv[i].lower() == '-noftr') :
            footerFlag = False
            nArguments -= 1

        elif (sys.argv[i].lower() == '-verbose') or \
                (sys.argv[i].lower() == '-v') :
            verboseFlag = True
            nArguments -= 1

        elif (sys.argv[i].lower() == '-help') or \
                (sys.argv[i].lower() == '-h') :
            helpFlag = True
            nArguments -= 1

        # unknown switch
        elif (sys.argv[i][0] == '-') :
            print(' ')
            print('ERROR: switch ' + sys.argv[i].upper() + ' not found')
            print('    Try ''./docx2text.py -help'' for more options')
            print(' ')
            exit()
            
    # Check to see if the minimum number of arguements (2) has been
    # supplied. Note that you don't need two arguments if the help
    # flag has been thrown
    if (helpFlag == False) and (nArguments != 2) :
        print(' ')
        print('ERROR: provide input and output filenames')
        print(' ')
        exit()

    # extract the filenames
    fileNameInput = sys.argv[-2]
    fileNameOutput = sys.argv[-1]

    # check to see if the specified input file exists
    if ( os.path.isfile(fileNameInput) == False ) :
        print (' ')
        print ('ERROR: input file ' + fileNameInput + ' not found')
        print (' ')
        exit()

    # Convert the word docx to text
    if helpFlag == False :

        if verboseFlag == True : print(' Opening input file ' + fileNameInput)

        # open the output file
        newfile = open(fileNameOutput, 'w')

        if headerFlag :
            # read the header (if requested)
            # note that there may be up to three header files
            # depending on odd/even/both, so we should check all three
            # to be safe. Status is true if text is found in any of them.
            if verboseFlag == True : print ' Searching for header ... ',
            status1 = getTheText(fileNameInput,newfile,'hdr1')
            status2 = getTheText(fileNameInput,newfile,'hdr2')
            status3 = getTheText(fileNameInput,newfile,'hdr3')
            status = status1 or status2 or status3
            if verboseFlag == True : print status

        if bodyFlag :
            # read the body (always requested)
            if verboseFlag == True : print ' Searching for body   ... ',
            status = getTheText(fileNameInput,newfile,'body')
            if verboseFlag == True : print status

        if footerFlag :
            # read the footer (if requested)
            if verboseFlag == True : print ' Searching for footer ... ',
            status1 = getTheText(fileNameInput,newfile,'ftr1')
            status2 = getTheText(fileNameInput,newfile,'ftr2')
            status3 = getTheText(fileNameInput,newfile,'ftr3')
            status = status1 or status2 or status3
            if verboseFlag == True : print status

        newfile.close()

    # if the user requests help, print the help screen
    else : 
        print(' ')
        print('DOCX2TXT.py : coverts an MS Word docx file to text')
        print('    ./docx2txt.py inputfile.docx outputfile.txt')
        print('    optional switches: -noheader (-nohdr), -nofooter (-noftr)')
        print('                       -verbose (-v), -help')
        print(' ')
    
    # end of main


def getTheText(fileNameInput,newfile,fileType):
    # This is the functiont that acutally opens the respective xml file
    # and reads and converts the text

    status = ' found'

    try :

        # open the respective xml file
        document = opendocx(fileNameInput,fileType)

        # extract the text from the xml file
        paratextlist = getdocumenttext(document)

        # if any text is found, make it unicode and write it to file

        if len(paratextlist) > 0 :
            # Make explicit unicode version
            newparatextlist = []
            for paratext in paratextlist:
                newparatextlist.append(paratext.encode("utf-8")+'\n')

            # Write the text to file
            newfile.write(''.join(newparatextlist)+'\n\n')

    except :

        # if the xml file isn't found
        status = ' not found'

    return status

    # end of getTheText

if __name__ == '__main__':
    main()
