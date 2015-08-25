#imports from pdfminer and others

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
from collections import defaultdict
import re
import os
import hashlib
import time
import logging

#defines the sort order for idenfitying precidence of revisions
#used to order revisions in rev_list when find_filename is called on a PDF
sort_alphabet = '-RrXx9876543210'

#starts the files processed count at zero
files_processed = 0

#defines the maximum number of seconds (in either direction) which will associate igs and dxf files with a pdf
#currently set to 30 minutes
association_window = 1800

#full paths to the various folders used to locate the files

#directory_path points to the directory crawled to find all the files--should point to vendor files
directory_path = os.path.normpath("X:\Engineering\Vendor Files")
#path to the archived drawings
archive_path = os.path.normpath("X:\Engineering\Archive")
#path to the file where drawing names which don't match the format will be moved
riffraff_path = os.path.normpath("X:\Engineering\Archive")

#test commit
#path to where the logs will be stored
logging.basicConfig(filename="X:\Engineering\Project Folders\Project 647 - Vendor Files Script\logs\log %s.txt"% time.strftime("%d-%b-%y %H%M"), level=logging.INFO)

try:
    os.remove("X:\Engineering\Project Folders\Project 647 - Vendor Files Script\In Progress QTD list.txt")
    
except:
    print("Couldn't delete QTD list")

try:
    qtdfile = open("X:\Engineering\Project Folders\Project 647 - Vendor Files Script\In Progress QTD list.txt", 'a')

except:
    print("Couldn't open new QTD list")
#------------------------------------------------convert_pdf_to_txt------------------------------------------------------

#Converts the entire pdf to a string of text, and returns this text. Needs to be passed a full path

def convert_pdf_to_txt(path):
    logging.debug('convert_pdf_to_txt called with:\n' + path + '\n')
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    maxpages = 0
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password="",caching=caching, check_extractable=True):
        interpreter.process_page(page)
    fp.close()
    device.close()
    str = retstr.getvalue()
    retstr.close()
    logging.debug('convert_pdf_to_txt returns the text of the PDF\n\n')
    return str

#------------------------------------------------find_filename-----------------------------------------------------------

#works only for pdf files

def find_filename(filepath):
    
    logging.debug('find_filename called with:\n' + filepath + '\n')

    if re.search("\d\d\d\d\d\d [RrXx]-\d\d\.",filepath) is not None:
#       comment out these lines to stop examination of each pdf and trust the name
#        try:
#            pdftext = convert_pdf_to_txt(filepath)
#            if re.search("QTD", pdftext) is not None:
#                print("yes")
#                qtdfile.write(filepath + "\n")
#        except:
#            pass
        try:
            pdftext = convert_pdf_to_txt(filepath)
            if re.search("QTD", pdftext) is not None:
                print("yes")
                qtdfile.write(filepath + "\n")
        except:
            pass
           
        rev = re.search("\d\d\d\d\d\d ([RrXx]-\d\d)\.",filepath).group(1)
        filenumber = re.search("\d\d\d\d\d\d",filepath).group()
        if rev!= None:
            if rev not in rev_list[filenumber]:
                rev_list[filenumber].append(rev)
                rev_list[filenumber] = sorted(rev_list[filenumber], key = lambda word: [sort_alphabet.index(c) for c in word])
        return os.path.basename(filepath)

    if file_to_open.lower().endswith('.pdf'):  
        pass
    else:
        logging.debug('main module: \n' + filepath + ' isn\'t a PDF \n\n')
        return os.path.basename(filepath)
    
    try:
        filenumber = re.search("\d\d\d\d\d\d",filepath).group()

        rev = find_rev_pdf(filepath)
        
        if rev!= None:
            if rev not in rev_list[filenumber]:
                rev_list[filenumber].append(rev)
                rev_list[filenumber] = sorted(rev_list[filenumber], key = lambda word: [sort_alphabet.index(c) for c in word])
    except:
        logging.debug('Could not find a 6 digit number in filename. Return:\n' + os.path.basename(filepath) + '\n\n')
        return os.path.basename(filepath)
    
    if rev != None:
        new_filename = filenumber + ' ' + rev + '.PDF'
        logging.debug('find_filename returns:\n' + new_filename + '\n\n')
        return new_filename
    else:
        return os.path.basename(filepath)

#------------------------------------------------non_pdf_find_filename---------------------------------------------------

#works only for igs and dxf files

def non_pdf_find_filename(filepath):

    logging.debug('non_pdf_find_filename called with:\n' + filepath + '\n') 
    
    try:
        filenumber = re.search("\d\d\d\d\d\d",filepath).group()
        rev = non_pdf_find_rev(filepath, filenumber)
    except:
        logging.debug('Could not find a 6 digit number in filename. Return:\n' + os.path.basename(filepath) + '\n\n')
        return os.path.basename(filepath)                

    if rev!= None:
        filename = re.search("(.+)(\..+)",os.path.basename(filepath))
        new_filename = filenumber+' '+rev+filename.group(2)
        return new_filename
    else:
        return os.path.basename(filepath)
    

#------------------------------------------------non_pdf_find_rev--------------------------------------------------------

#works only for igs and dxf files
    
def non_pdf_find_rev(filepath, filenumber):

    bottom = os.path.getmtime(filepath) - association_window
    top = os.path.getmtime(filepath) + association_window

    for revision in rev_list[filenumber]:
        pdf_path = os.path.join(directory_path, filenumber+' '+revision+'.pdf')
        pdf_time = os.path.getmtime(pdf_path)
        if (bottom < pdf_time and pdf_time < top):
            return revision
    return

#------------------------------------------------find_rev_pdf------------------------------------------------------------

#Returns rev information as a string of undermined length if passed a valid file ending in ".pdf"
#Returns None if passed an invalid filename and explains that the file doesn't exist
#Silently returns None if passed a file that doesn't end in ".pdf"

def find_rev_pdf(filepath):
    logging.debug('find_rev_pdf called with:\n' + filepath + '\n')
    if filepath[-4:].lower() == ".pdf":
        try:
            pdftext = convert_pdf_to_txt(filepath)
            revinfo = re.search("REV #\n(\S*\d)",pdftext)
            logging.debug('find_rev_pdf returns:\n'+ revinfo.group(1) + '\n\n')
            revision_on_dwg = revinfo.group(1)

            if re.search("^\d$",revision_on_dwg) is not None:
                return "R-0" + str(revision_on_dwg)
            else:
                return revision_on_dwg
        except:
            logging.debug('find_rev_pdf returns:\n'+ filepath + '\n File didn\'t match regex\n\n')
            return
    else:
        logging.debug('find_rev_pdf was called on a file that is not a PDF.\n')
        return

#------------------------------------------------list_files--------------------------------------------------------------

#when passed a directory path, lists all the files in the folder, and all subfolders except other
#Returns full paths to each file

def list_files(directory_path):
    logging.debug('list_files called with:\n' + directory_path + '\n')
    filenames_list = []
    for dirpath, dirnames, filenames in os.walk(directory_path):
        if riffraff_path in dirpath:
            dirpath.remove(riffraff_path)
        filenames_list.extend(os.path.join(dirpath,x) for x in filenames)
    logging.debug('list_files returns file list\n\n')
    return filenames_list

#------------------------------------------------gracefulRename----------------------------------------------------------

#renames the first filepath called to the second filepath called.
#if the filepaths are the same, no change is made
#if the second filepath already exists, the files are both hashed, if they are the same, the original_filepath file is deleted
#if the hashes don't match, the filename will be modified until it doesn't collide
#depends on extension being 4 characters (including period)

def gracefulRename(original_filepath, new_filepath):
    logging.debug('gracefulRename called with:\n' + original_filepath + '\n' + new_filepath + '\n')
    if original_filepath.lower() == new_filepath.lower():
        logging.debug('gracefulRename attempted to rename a file to the same name. No change was made\n\n')
        return
    if os.path.exists(new_filepath):
        logging.debug('gracefulRename is attempting to rename a file to a path that already exists\n')
        if hashfile(original_filepath) == hashfile(new_filepath):
            os.remove(original_filepath)
            logging.info('gracefulRename returns: ' + os.path.basename(original_filepath) + ' and ' + os.path.basename(new_filepath) + ' were identical, removed ' + original_filepath +'\n\n')
            return
        else:
            attempts = 1
            while True:
                modified_new_filepath = new_filepath[:-4] + ' (%i)' % attempts + new_filepath[-4:]
                if os.path.exists(modified_new_filepath):
                    attempts += 1
                else:
                    try:
                        os.rename(original_filepath, modified_new_filepath)
                        logging.info('gracefulRename returns: \n' + original_filepath + '\n renamed to \n' + modified_new_filepath + '\n\n')
                        return
                    except:
                        return
                
    else:
        try:
            os.rename(original_filepath, new_filepath)
            logging.info(' gracefulRename changed ' + os.path.basename(original_filepath) + ' to ' + os.path.basename(new_filepath) + ' successfully\n\n')
        except:
            return

#------------------------------------------------hashfile----------------------------------------------------------------
        
#when passed a filepath, calculates the md5 hash and returns it as a string

def hashfile(filepath):
    try:
        logging.debug('hashfile called with: \n' + filepath + '\n')    
        afile = open(filepath, 'rb')
        hasher = hashlib.md5()
        buf = afile.read(65536)
        while len(buf) > 0:
            hasher.update(buf)
            buf = afile.read(65536)
        afile.close()
        logging.debug('hashfile returns: ' + hasher.hexdigest() + '\n\n')
        return hasher.hexdigest()
    except:
        return 1

#------------------------------------------------main loop---------------------------------------------------------------


#call file list to return a list of full paths to all files in vendor files
file_list = list_files(directory_path)

#initialize the dict for storing revlist
rev_list = defaultdict(list)

#loop through each file in the file list
for file_to_open in file_list:
    files_processed += 1
    print files_processed
    #log the interaction with the file under this heading
    logging.debug('\n--------------------------------------- First Pass ' + os.path.basename(file_to_open) + ' ---------------------------------------\n\n')

    if re.search("(\d\d\d\d\d\d [RrXx]-\d\d\.|\d\d\d\d\d\d\.|\d\d\d\d\d\d [RrXx]\d+.|\d\d\d\d\d\d \d+\.|\d\d\d\d\d\d +\.|\d\d\d\d\d\d [RrXx]\d+\.|\d\d\d\d\d\d REV.+\.)",os.path.basename(file_to_open)) is not None:
        #file has the correct name format, if it's a pdf, it can be renamed now, and it's rev added to revlist 
        gracefulRename(file_to_open, os.path.join(directory_path, find_filename(file_to_open)))

    else:
        #file doesn't have the correct name format. needs to be moved to riffraff
        logging.info('main module: \n' + file_to_open + ' doesn\'t match the name format, moved to archive \n\n')
        gracefulRename(file_to_open, os.path.join(archive_path, os.path.basename(file_to_open)))
        
    
logging.debug('First pass finished\n\n')

#rev_list is complete and ordered at this point

qtdfile.close()

try:
    os.remove("X:\Engineering\Project Folders\Project 647 - Vendor Files Script\QTD list.txt")
    
except:
    print("Couldn't delete QTD list")

os.rename("X:\Engineering\Project Folders\Project 647 - Vendor Files Script\In Progress QTD list.txt","X:\Engineering\Project Folders\Project 647 - Vendor Files Script\QTD list.txt")

file_list = list_files(directory_path)

for file_to_open in file_list:

    files_processed += 1
    print files_processed
    
    logging.debug('\n--------------------------------------- Second Pass ' + os.path.basename(file_to_open) + ' ---------------------------------------\n\n')

    new_filepath = os.path.join(directory_path, non_pdf_find_filename(file_to_open))
    gracefulRename(file_to_open, new_filepath)

    if re.search("(\d\d\d\d\d\d .+\d\.|\d\d\d\d\d\d \d+\.)",os.path.basename(new_filepath)) is not None:
        #move old revs to archive here
        rev_info = re.search("(\d\d\d\d\d\d).+(\..+)",os.path.basename(new_filepath))
        extension = rev_info.group(2)
        number = rev_info.group(1)
        number_of_revs = len(rev_list[number])
        if number_of_revs > 1:
            for rev in rev_list[number][1:]:             
                gracefulRename(os.path.join(os.path.dirname(new_filepath), number + ' ' + rev + extension), os.path.join(archive_path, number + ' ' + rev + extension))            
        
    else:
        logging.info('main module: \n' + new_filepath + ' couldn\'t be assigned a revision, moved to archive \n\n')
        gracefulRename(new_filepath, os.path.join(archive_path, non_pdf_find_filename(new_filepath)))
