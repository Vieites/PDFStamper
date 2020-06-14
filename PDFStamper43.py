""""
PDFStamper v4.3  - Codename: 'Alan Southall'
An automated Python 3.7 script to stamp PDFs with a watermark and batch numbers
(c) Carlos Vieites - 2020 - All rights reserved.
"""
# Import system libraries
import os
import sys
import io
import time
import subprocess
from pathlib import Path
from keyring import get_keyring
get_keyring()
from getpass import getuser
import keyring
import win32timezone
from keyring.backends import Windows
import msvcrt
keyring.set_keyring(Windows.WinVaultKeyring())

# Define functions --------------------------------------------------------------------------------------------
# Download a PDF live document from MasterControl - autentication by cookie
def download_file(mcdocid, username, password):
    #Define REST API MasterControl URLs
    mcrestapi = 'https://btgplc.mastercontrol.com/btgplc/restapi/v1/'
    mcurl ='https://btgplc.mastercontrol.com/btgplc/restapi/identity/authentication/login?action=login'
    #Open session in MC with cookie saved for future requests
    s = requests.Session()
    login_data =  {'username':username, 'password':password}
    s.post(mcurl, login_data)
    #Retrieve Infocard metadata MC document ID
    infocardurl = mcrestapi + 'document/' + mcdocid + '/released-revision'
    response = s.get(infocardurl)
    infocardmetadata = response.json()
    infocardid = (infocardmetadata['infocardId'])
    #Diferenciate between Form and BMR by adding -Form or -BR to the output name
    doctype = (infocardmetadata['infocardTypeName'])
    if doctype == 'Form':
        outputfile = mcdocid + '-Form.PDF'
    else:
        outputfile = mcdocid + '-BR.PDF'
    #Download PDF file
    url = mcrestapi + 'document/' + infocardid + '/publishedMainFile'
    fileout = temppath + outputfile
    response = s.get(url)
    filename = Path(fileout)
    filename.write_bytes(response.content)
    time.sleep(0.5)
    return outputfile


# Download an SI from Sharepoint given a partial name, saves the file and returns filename
def download_si(si_partialname, bscusername, bscpassword):
    import shutil
    import contextlib
    usernamesi = bscusername + '@bsci.com'
    import sharepy
    import json
    server_url = 'https://bostonscientific.sharepoint.com'
    site_url = '/sites/Farnham-QA'
    library = 'Shared%20Documents/Document%20Control/Documentation%20Centre/Special%20Instructions'
    # Authenticate in the Office 365 sharepoint server, supress msgs with contextlib
    with contextlib.redirect_stdout(None):
        s = sharepy.connect(server_url, usernamesi, bscpassword)
        s.save()
    # List all SI files in the sharepoint dedicated library - output list as XML sharepoint.json
    office365_url_command = server_url + site_url + "/_api/Web//GetFolderByServerRelativeUrl('" + library + "')/Files?$expand=File&$filter=startswith(Name,'" + si_partialname + "')"
    r = s.get(office365_url_command)
    data = r.json()
    file = open("sharepoint.json", "w")
    file.write(json.dumps(data, indent=4))
    # Interrogate json list to get SI details (*** asuming unique file ***)
    si_filename = data['d']['results'][0]['Name']
    si_time_created = data['d']['results'][0]['TimeCreated']
    si_relativeurl = data['d']['results'][0]['ServerRelativeUrl']
    # Download SI
    si_url = (server_url + si_relativeurl)
    r = s.getfile(si_url, filename=si_filename)
    return si_filename


# Inject batch number in already watermarked blank Portrait and Landscape PDF page
def insertbn(font_type, font_size, xp, yp, xl, yl, batch):
    # Portrait
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont(font_type, font_size)
    can.drawString(xp, yp, batch)
    can.save()
    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read the existing PDF
    existing_pdf = PdfFileReader(open("BlankStampP.pdf", "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open("StampP.pdf", "wb")
    output.write(outputStream)
    outputStream.close()
    # Landscape
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont(font_type, font_size)
    can.drawString(xl, yl, batch)
    can.save()
    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read the existing PDF
    existing_pdf = PdfFileReader(open("BlankStampL.pdf", "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open("StampL.pdf", "wb")
    result = output.write(outputStream)
    outputStream.close()
    return result


# Stamper engine (merge numbered blank PDF stamps with decrypted inputed PDF)
def Stamper(infile, outfile):
    inputfile = open(infile, 'rb')
    pdfReader = PyPDF2.PdfFileReader(inputfile)
    pdfWriter = PyPDF2.PdfFileWriter()
    for pageNum in range(pdfReader.numPages):
        inputfilePage = pdfReader.getPage(pageNum)
        page = pdfReader.getPage(pageNum).mediaBox
        if (page.getUpperRight_x() - page.getUpperLeft_x()) > (page.getUpperRight_y() - page.getLowerRight_y()):
            stamp = 'StampL.pdf'
        else:
            stamp = 'StampP.pdf'
        pdfWatermarkReader = PyPDF2.PdfFileReader(open(stamp, 'rb'))
        inputfilePage.mergePage(pdfWatermarkReader.getPage(0))
        inputfilePage.compressContentStreams()
        pdfWriter.addPage(inputfilePage)
    resultPdfFile = open(outfile, 'wb')
    pdfWriter.write(resultPdfFile)
    inputfile.close()
    resultPdfFile.close()
    return resultPdfFile


# Print function with Duplex selection
def printpdf(infile, copies, setduplex):
    # duplex, dupexshort, duplexlong and simplex.
    printsettings = '"' + setduplex + ',' + str(copies) + '"x "'
    args = ' -print-to-default -exit-on-print -silent -print-settings ' + printsettings, infile
    result = subprocess.Popen(["SumatraPDF.exe", args], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
    time.sleep(2)
    result.stdin.close()
    return result


# Timer: waits timeout seconds and return timeflag = 1 if a key is pressed
def timer(timeout, message):
    timeflag = 0
    startTime = time.time()
    inp = None
    print('\n')
    print(message)
    print('\n')
    while True:
        if msvcrt.kbhit():
            inp = msvcrt.getch()
            break
        elif time.time() - startTime > timeout:
            break
    if inp:
        timeflag = 1
    else:
        timeflag = 0
    return timeflag


# Password control: request the username and password and store them securely
def userdetails(systemname):
    if systemname == 'btgpdfstamper':
        windowtitle = 'BTG Network Credentials'
        msg = 'Enter your BTG credentials to access the BTG iQMS MasterControl (e.g. asouthall): '
        fieldnames = ["BTG Username", "BTG Password"]
    else:
        systemname = 'bscpdfstamper'
        windowtitle = 'BSC Network Credentials'
        msg = 'Enter your BSC credentials to access the BSC SharePoint (e.g. smithj): '
        fieldnames = ["BSC Username", "BSC Password"]
    userfile = systemname + '.pkl'
    from easygui import multpasswordbox
    fieldvalues = []
    fieldvalues = multpasswordbox(msg, windowtitle, fieldnames)
    username = fieldvalues[0]
    userpassword = fieldvalues[1]
    f = open(userfile, 'wb')
    pickle.dump(username, f)
    f.close()
    keyring.set_password(systemname, username, userpassword)
    return fieldvalues


# Error handling: displays a message box and exit the script
def error_handling(msg):
    from easygui import msgbox
    title = "Script Error !!"
    ok_button = 'Exit'
    msgbox('Please resolve the following and re-start: ' + msg, title, ok_button)
    raise SystemExit


# Warning message: displays a warning message to the user
def warning_message(title, msg):
    choices = ["Exit", "Continue"]
    from easygui import buttonbox
    reply = buttonbox(msg, choices=choices, title=title)
    if reply == "Exit":
        raise SystemExit
    else:
        pass
# End of the define functions section ----------------------------------------------------------------------------


# Apply color to the console
os.system('color 71')

# Welcome message - display a splash screen
print('\n')
print('Welcome to PDFStamper 4.3 release - Codename: Alan Southall.')
print('\n')
import wx
import wx.lib.agw.advancedsplash as AS
app = wx.App(0)
imagePath = "splash43.png"
bitmap = wx.Bitmap(imagePath, wx.BITMAP_TYPE_PNG)
shadow = wx.BLUE
splash = AS.AdvancedSplash(None, bitmap=bitmap, timeout=2000, agwStyle=AS.AS_TIMEOUT | AS.AS_CENTER_ON_PARENT | AS.AS_SHADOW_BITMAP, shadowcolour=shadow)
app.MainLoop()

# Handle password to access BTG MasterControl and BSC SharePoint
# Handle BTG network username and password
import keyring
import win32timezone
import pickle
import msvcrt
systemname = 'btgpdfstamper'
userfile = systemname + '.pkl'
try:
    f = open(userfile, 'rb')
    btgusername = pickle.load(f)
    f.close()
    btgpassword = keyring.get_password(systemname, btgusername)
except:
    values = userdetails(systemname)
    btgusername = values[0]
    btgpassword = values[1]
# Handle BSC network username and password
systemname = 'bscpdfstamper'
userfile = systemname + '.pkl'
try:
    f = open(userfile, 'rb')
    bscusername = pickle.load(f)
    f.close()
    bscpassword = keyring.get_password(systemname, bscusername)
except:
    values = userdetails(systemname)
    bscusername = values[0]
    bscpassword = values[1]

# User to input file Excel file
print('\n')
print('Loading default request file "PDFStamper list.xlsx"... ')
timeflag = timer(2, 'Wait 3 seconds, or press any key now to change the request file.')
if timeflag == 1:
    from easygui import fileopenbox
    userfile = fileopenbox()
    print(userfile)
else:
    userfile = 'PDFStamper list.xlsx'

# Read user request excel file with 4 columns: Document, Batch Number, Prints, Duplex
import pandas as pd
inputlist = pd.read_excel(userfile, usecols='A,B,C,D')
df = pd.DataFrame(inputlist)
userlist = pd.DataFrame(inputlist)

# Download document pack codes list from BTG MasterControl
# MasterControl FAR-02336, Proquis FAR-SP-02.001-F11
mcurl ='https://btgplc.mastercontrol.com/btgplc/restapi/identity/authentication/login?action=login'
codesurl = 'https://btgplc.mastercontrol.com/btgplc/Main/mastercontrol/vault/view_doc.cfm?ls_id=B6YDBJUUYZBVVKEY7O&download=true'
outputfile = 'Codes.xlsx'
import msvcrt
import win32timezone
import requests
from requests import get
#Open session in MC with cookie saved for future requests
s = requests.Session()
login_data =  {'username':btgusername, 'password':btgpassword}
s.post(mcurl, login_data)
#Download codes spreadsheet from MC
try:
    response = s.get(codesurl)
    filename = Path(outputfile)
    filename.write_bytes(response.content)
except Exception:
    msg = 'Codes list spreadsheet is open. Close and restart.'
    error_handling(msg)

# Read Pack Codes excel spreadsheet file with 4 columns: Pack ID, MasterControl, Copies, Multiple
try:
    codeslist = pd.read_excel(outputfile, usecols='C,I,F,G')
    codes = pd.DataFrame(codeslist)
except Exception:
    msg = 'Unable to access BTG MasterControl to read the codes list. Restart to refresh the BTG password. (Warning: using the WRONG password 3 times will lock the user in MasterControl).'
    os.remove('btgpdfstamper.pkl')
    error_handling(msg)

# Use the 'Pack Codes' list to construct the working list to download
workinglist_id = []
workinglist_copies = []
workinglist_batch = []
workinglist_duplex = []
wrongcode = 0

for i in range(len(userlist['Document'])):
    # Error catching blank entries
    if type(userlist.loc[i, 'Document']) == float:
        msg = 'BLANK entry as a Document Number in line ' + str(i+2) + ' in the request list.'
        error_handling(msg)
    # Requested as a traditional FAR- document
    if 'FAR-' in (userlist.loc[i, 'Document']):
        workinglist_id.append(userlist.loc[i, 'Document'])
        workinglist_copies.append(userlist.loc[i, 'Copies'])
        workinglist_batch.append(userlist.loc[i, 'BatchNo'])
        workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
    # Special Instruction SI requested
    elif 'SI-' in (userlist.loc[i, 'Document']):
        workinglist_id.append(userlist.loc[i, 'Document'])
        workinglist_copies.append(userlist.loc[i, 'Copies'])
        workinglist_batch.append(userlist.loc[i, 'BatchNo'])
        workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
    else:
        # Requested as a document pack code
        wrongcode = 1
        for j in range(len(codes['Pack ID'])):
            if (userlist.loc[i, 'Document']) in (codes.loc[j, 'Pack ID']):
                wrongcode = 0
                workinglist_id.append(codes.loc[j, 'MasterControl'])
                workinglist_copies.append(codes.loc[j, 'Copies'])
                workinglist_batch.append(userlist.loc[i, 'BatchNo'])
                # Error catching: BMRs with multiple batch numbers are not yet supported
                if (codes.loc[j, 'Multiple']) == 'M':
                    multiplecodename = userlist.loc[i, 'Document']
                    multiplecodebmr = codes.loc[j, 'MasterControl']
                    msg = 'Requested code ' + multiplecodename + ' contains ' + multiplecodebmr + ' which uses MULTIPLE batch numbers and it is not yet supported. Continue and stamp anyway?'
                    warning_message('Multiple batch numbers!', msg)
                # All Forms to be printed double sided (duplex on)
                if '-Form' in (codes.loc[j, 'MasterControl']) and '-BR' not in (codes.loc[j, 'MasterControl']):
                    workinglist_duplex.append('D')
                else:
                    workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
    # Error catching a requested wrong code using flag wrongcode = 1
    if wrongcode == 1:
        msg = 'Document Pack Code ' + (userlist.loc[i, 'Document']) + ' is not valid. Check the request list.'
        error_handling(msg)

# Inform user showing the list of files the script will work on
import numpy as np
downloadlist = np.unique(workinglist_id)
print('Your request contains', len(workinglist_id), 'documents, of which: ')
print(str(len(downloadlist)), 'are unique (no repeats) to be retrieved from MasterControl or SharePoint: ')
print(downloadlist)

# Prep a temp folder for downloads
# Check if the temp folder exists otherwise, create it
temppath = 'C:/pdfstampertemp/'
if not os.path.exists(temppath):
    os.makedirs(temppath)
# Check temp folder is empty otherwise, clean it
if os.listdir(temppath) != []:
    fileList = os.listdir(temppath)
    for fileName in fileList:
        try:
            os.remove(temppath + fileName)
        except Exception:
            msg = 'Unable to clear temp folder. ' + fileName + ' is in use. Close it.'
            error_handling(msg)

# Download files from BTG MasterControl and BSC SharePoint
from tqdm import tqdm
print('\n')
print('DOWNLOADING FILES - ' + str(len(downloadlist)) + ' documents')
downloadlist2 = []
for x in tqdm(range(len(downloadlist)), unit='file'):
    # MasterControl document
    if 'FAR-' in str(downloadlist[x]):
        mcdocid = str(downloadlist[x])
        tqdm.write('Downloading: ' + mcdocid)
        result = download_file(mcdocid, btgusername, btgpassword)
        downloadlist2.append(str(result))
    # Special instruction from BSC Teams SharePoint
    elif 'SI-' in str(downloadlist[x]):
        tqdm.write('Downloading: ' + downloadlist[x])
        try:
            si_partialname = str(downloadlist[x])
            si_filename = download_si(si_partialname, bscusername, bscpassword)
        except Exception:
            msg = 'Unable to download ' + si_partialname + '. Remove it from the list and restart the script.'
            error_handling(msg)
        # Rename file to requested SI and move the downloaded SI to the temp folder
        import shutil
        os.rename(si_filename, si_partialname + '.PDF')
        shutil.move(si_partialname + '.PDF', temppath)

# Decrypt downloaded PDFs
print('\n')
print('PREPARING PDFs - ' + str(len(downloadlist2)) + ' documents')
# Call qpdf to decrypt pdfs
for x in tqdm(range(len(downloadlist2)), unit='PDF'):
    tqdm.write('Preparing ' + downloadlist2[x])
    inputfile = temppath + downloadlist2[x]
    outputfile = temppath + 'D' + downloadlist2[x]
    subprocess.run(["qpdf.exe", "--decrypt", inputfile, outputfile])

# Stamp decrypted PDFs with batch numbers
print('\n')
print('STAMPING PDFs - ' + str(len(workinglist_id)) + ' documents')
print('This may take a while according with the size of the PDF. Please wait...')
# Prep: declare variables to be used to create stamps form blank watermarked
stamp = ['StampP.pdf', 'StampL.pdf']
font_type = 'Times-Bold'
font_size = 16

# Upadate the workinglist with the info extracted during download to differentiate BMR and Forms
for x in range(len(workinglist_id)):
    for y in range(len(downloadlist2)):
        if workinglist_id[x] in downloadlist2[y]:
            workinglist_id[x] = downloadlist2[y]

# Stamp PDF
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
looprange = len(workinglist_id)
for i in tqdm(range(looprange), unit='PDF'):
    tqdm.write(workinglist_id[i])
    inputpdf = temppath + 'D' + workinglist_id[i]
    outputpdf = temppath + str(i) + 'SD' + workinglist_id[i]
    if '-BR' in inputpdf:
        currentdoctype = 'BR'
        xp = 420
        yp = 765
        xl = 645
        yl = 520
    else:
        currentdoctype = 'F'
        xp = 420
        yp = 780
        xl = 640
        yl = 530
    insertbn(font_type, font_size, xp, yp, xl, yl, str(workinglist_batch[i]))
    # Call the function to do the stamping
    if 'SI-' in inputpdf:
        try:
            args = " --overlay StampP.pdf --repeat=1-z -- " + inputpdf, " " + outputpdf
            subprocess.run(["qpdf", args])
        except Exception:
            msg = 'Unable to stamp ' + workinglist_id[i] + '. Remove it from the request list and re-start.'
            error_handling(msg)
    else:
        try:
            Stamper(inputpdf, outputpdf)
        except Exception:
            msg = (workinglist_id[i]) + '  1.- Is it a PDF (not Word)?   2.- Incorrect document ID number?  3.-Locked in MasterControl?'
            error_handling(msg)
print('\n')
print("Stamping has completed!")
print('\n')

# Send to the printer
print('PRINTING - ' + str(len(workinglist_id)) + ' documents')
for x in tqdm(range(looprange), unit='PDF'):
    if workinglist_duplex[x] == 'D' or workinglist_duplex[x] == 'd':
        setduplex = 'duplex'
    else:
        setduplex = 'simplex'
    filetoprint = temppath + str(x) + 'SD' + str(workinglist_id[x])
    nocopies = int(workinglist_copies[x])
    tqdm.write(filetoprint)
    printpdf(filetoprint, nocopies, setduplex)
    # Give some time for the file to reach the printing buffer
    time.sleep(4)

# Goodbye message
print('\n')
print('Documents sent to the printer.')
print('All done. - Script terminated. Have a nice day!')
print('\n')

# Ask user to open folder with resulting stamped files
msg = 'All done. - ' + str(len(workinglist_id)) + ' documents sent to the printer. Have a nice day!'
choices = ["Open Temp folder", "Exit"]
from easygui import buttonbox
reply = buttonbox(msg, choices=choices, title='Operation Completed')
if reply == "Open Temp folder":
    path = os.path.realpath(temppath)
    os.startfile(path)
raise SystemExit
