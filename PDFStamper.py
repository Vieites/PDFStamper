"""
PDFStamper v4.1
An automated Python 3.7 script to stamp PDFs with a watermark and batch numbers
(c) Carlos Vieites - 2019 - All rights reserved.
"""
# Import system libraries
import os
import sys
import io
import time
import subprocess
from pathlib import Path

# Define functions --------------------------------------------------------------------------------------------
# Download a file from https Proquis with authentication handled by getpass and keyring
def download_file(url, outputfile, username, password):
    response = get(url, allow_redirects=True, auth=HttpNtlmAuth(username, password))
    filename = Path(outputfile)
    filename.write_bytes(response.content)
    time.sleep(0.5)
    return outputfile


# Download an SI from Sharepoint given a partial name, saves the file and returns filename
def download_si(si_partialname, username, userpassword):
    import shutil
    import contextlib
    usernamesi = username + '@btgplc.com'
    import sharepy
    import json
    server_url = 'https://btggroup.sharepoint.com'
    site_url = '/sites/QA-Farnham'
    library = 'Shared%20Documents/Document%20Control/Documentation%20Centre/Special%20Instructions'
    # Authenticate in the Office 365 sharepoint server, supress msgs with contextlib
    with contextlib.redirect_stdout(None):
        s = sharepy.connect(server_url, usernamesi, userpassword)
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


# Password control - ask user for the Proquis password
def passwordcontrol(username):
    systemname = "PDFStamper"
    from easygui import passwordbox
    userpassword = passwordbox('Password to access Proquis:')
    keyring.set_password(systemname, username, userpassword)
    userpassword = keyring.get_password(systemname, username)
    return userpassword


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
print('Welcome to PDFStamper!')
print('\n')
import wx
import wx.lib.agw.advancedsplash as AS
app = wx.App(0)
imagePath = "splash.png"
bitmap = wx.Bitmap(imagePath, wx.BITMAP_TYPE_PNG)
shadow = wx.BLUE
splash = AS.AdvancedSplash(None, bitmap=bitmap, timeout=2000, agwStyle=AS.AS_TIMEOUT | AS.AS_CENTER_ON_PARENT | AS.AS_SHADOW_BITMAP, shadowcolour=shadow)
app.MainLoop()

# Check for updates, if new update is found, ask user to download and install
print('Checking for new version...', end=' ')
import csv
import urllib.request
current_version = 4.1
update_flag = 0
# Check the web for a csv file indicating the current released version
new_version_check_url = 'https://onedrive.live.com/download?cid=D7B6F839EC0FBFD5&resid=D7B6F839EC0FBFD5%2133714&authkey=AJUv0OrRjes1zbA'
try:
    webpage = urllib.request.urlopen(new_version_check_url)
    reader = csv.reader(io.TextIOWrapper(webpage))
    new_version = []
    for line in reader:
        new_version = line
    # Compare running version against web reported version and ask user to update if newer
    if (float(new_version[0]) > float(current_version)):
        print('New version found, PDFStamper version', new_version[0])
        update_flag = 1
        msg = 'New version v' + new_version[0] + ' has been released!! Would you like to UPDATE to the new version now?'
        choices = ["Ignore", "Update Now"]
        from easygui import buttonbox
        reply = buttonbox(msg, choices=choices, title='New Version Available!')
        if reply == "Update Now":
            print('\n')
            print('DOWNLOADING UPDATE. Please wait...')
            from requests import get
            with open('PDFStamper Setup.exe', 'wb') as file:
                from tqdm import tqdm
                for i in tqdm(range(100), unit=' bytes'):
                    response = get(new_version[1])
                    file.write(response.content)
                    args = 'PDFStamper Setup.exe'
        # User has ignored the update
        elif reply == "Ignore":
            update_flag = 0
    else:
        print('You are up to date!')
        update_flag = 0
except Exception:
    update_flag = 0
    print('Unable to check for new updates.')
# If user accepted the update now, execute the downloaded update file and terminate script
if update_flag == 1:
    subprocess.Popen(args)
    sys.exit(0)

# Handle password to access HTTPS Proquis and SharePoint
from getpass import getuser
import keyring
import win32timezone
from keyring.backends import Windows
import msvcrt
keyring.set_keyring(Windows.WinVaultKeyring())
username = getuser()
userpassword = keyring.get_password("PDFStamper", username)
if (userpassword is None):
    userpassword = passwordcontrol(username)

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

# Download document pack codes list from Proquis FAR-SP-02.001-F11
from requests import get
from requests_ntlm import HttpNtlmAuth
codesurl = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO=FAR-SP-02.001-F11'
outputfile = 'Codes.xlsx'
try:
    download_file(codesurl, outputfile, username, userpassword)
except Exception:
    msg = '"Codes.xlsx" is open. Close the Excel file and re-start.'
    error_handling(msg)

# Read codes excel file with 4 columns: Pack ID, Document, Copies, Multiple
erroraccessingproquis = 1
while erroraccessingproquis == 1:
    try:
        codeslist = pd.read_excel(outputfile, usecols='E,G,H,I')
        codes = pd.DataFrame(codeslist)
        erroraccessingproquis = 0
    except Exception:
        from easygui import ccbox
        title = 'Error Accessing Proquis'
        msg = 'Unable to read the "Pack Codes" list from Proquis. Continue and re-enter BTG Password?'
        response = ccbox(msg, title)
        if response is True:
            userpassword = passwordcontrol(username)
            download_file(codesurl, outputfile, 'intl\\' + username, userpassword)
        else:
            raise SystemExit

# Use the 'Pack Codes' list to construct the working list to download
workinglist_id = []
workinglist_copies = []
workinglist_batch = []
workinglist_duplex = []
wrongcode = 0

for i in range(len(userlist['Document'])):
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
                workinglist_id.append(codes.loc[j, 'Document'])
                workinglist_copies.append(codes.loc[j, 'Copies'])
                workinglist_batch.append(userlist.loc[i, 'BatchNo'])
                # Error catching: BMRs with multiple batch numbers are not yet supported
                if (codes.loc[j, 'Multiple']) == 'M':
                    multiplecodename = userlist.loc[i, 'Document']
                    multiplecodebmr = codes.loc[j, 'Document']
                    msg = 'Requested code ' + multiplecodename + ' contains ' + multiplecodebmr + ' which uses MULTIPLE batch numbers and it is not yet supported. Continue and stamp anyway?'
                    warning_message('Multiple batch numbers!', msg)
                # All Forms to be printed double sided (duplex on)
                if '-F' in (codes.loc[j, 'Document']) and '-BR' not in (codes.loc[j, 'Document']):
                    workinglist_duplex.append('D')
                else:
                    workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
    # Error catching a requested wrong code using flag wrongcode = 1
    if wrongcode == 1:
        msg = 'Document Pack Code ' + (userlist.loc[i, 'Document']) + ' is not valid. Check request list.'
        error_handling(msg)

# Inform user showing the list of files the script will work on
import numpy as np
downloadlist = np.unique(workinglist_id)
print('Your request contains', len(workinglist_id), 'documents, of which: ')
print(str(len(downloadlist)), 'are unique (no repeats) to be retrieved from Proquis and/or SharePoint: ')
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

# Download files from SSL Proquis and SharePoint
from tqdm import tqdm
print('\n')
print('DOWNLOADING FILES - ' + str(len(downloadlist)) + ' documents')
url = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO='
for x in tqdm(range(len(downloadlist)), unit='file'):
    if 'FAR-' in str(downloadlist[x]):
        urlfile = url + str(downloadlist[x])
        outputfile = temppath + str(downloadlist[x]) + '.PDF'
        tqdm.write('Downloading: ' + urlfile)
        download_file(urlfile, outputfile, username, userpassword)
    elif 'SI-' in str(downloadlist[x]):
        tqdm.write('Downloading: ' + downloadlist[x])
        try:
            si_partialname = str(downloadlist[x])
            si_filename = download_si(si_partialname, username, userpassword)
        except Exception:
            msg = 'Unable to download ' + si_partialname + '. Remove it from the list and restart the script.'
            error_handling(msg)
        # Rename file to requested SI and move the downloaded SI to the temp folder
        import shutil
        os.rename(si_filename, si_partialname+'.PDF')
        shutil.move(si_partialname+'.PDF', temppath)

# Decrypt downloaded PDFs
print('\n')
print('PREPARING PDFs - ' + str(len(downloadlist)) + ' documents')
# Call qpdf to decrypt pdfs
for x in tqdm(range(len(downloadlist)), unit='PDF'):
    tqdm.write('Preparing ' + downloadlist[x])
    inputfile = temppath + downloadlist[x] + '.PDF'
    outputfile = temppath + 'D' + downloadlist[x] + '.PDF'
    subprocess.run(["qpdf.exe", "--decrypt", inputfile, outputfile])

# Stamp decrypted PDFs with batch numbers
print('\n')
print('STAMPING PDFs - ' + str(len(workinglist_id)) + ' documents')
print('This may take a while according with the size of the PDF. Please wait...')
# Prep: declare variables to be used to create stamps form blank watermarked
stamp = ['StampP.pdf', 'StampL.pdf']
font_type = 'Times-Bold'
font_size = 16

# Stamp PDF
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
looprange = len(workinglist_id)
for i in tqdm(range(looprange), unit='PDF'):
    tqdm.write(workinglist_id[i])
    inputpdf = temppath + 'D' + workinglist_id[i] + '.PDF'
    outputpdf = temppath + str(i) + 'SD' + workinglist_id[i] + '.PDF'
    if '-BR-' in inputpdf:
        currentdoctype = 'BR'
        xp = 425
        yp = 770
        xl = 645
        yl = 520
    else:
        currentdoctype = 'F'
        xp = 430
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
            msg = (workinglist_id[i]) + '  1.- Is it a PDF (not Word)?   2.- Incorrect document ID number?  3.-Locked in Proquis?'
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
    filetoprint = temppath + str(x) + 'SD' + str(workinglist_id[x]) + '.PDF'
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
