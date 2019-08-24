"""
PDFStamper v3.5
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
# Download a file from https with authentication handled by getpass and keyring
def download_file(url, outputfile, username, password):
    response = get(url, allow_redirects=True, auth=HttpNtlmAuth(username, password))
    filename = Path(outputfile)
    filename.write_bytes(response.content)
    time.sleep(0.5)
    return outputfile


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
    #
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

# End of the define functions section ----------------------------------------------------------------------------


# Welcome message - display a splash screen
import wx
import wx.lib.agw.advancedsplash as AS
app = wx.App(0)
imagePath = "splash.png"
bitmap = wx.Bitmap(imagePath, wx.BITMAP_TYPE_PNG)
shadow = wx.BLUE
splash = AS.AdvancedSplash(None, bitmap=bitmap, timeout=2000, agwStyle=AS.AS_TIMEOUT | AS.AS_CENTER_ON_PARENT | AS.AS_SHADOW_BITMAP, shadowcolour=shadow)
app.MainLoop()

# Apply color to the console
os.system('color 71')

# Handle password to access HTTPS Proquis
from getpass import getuser
import keyring
import win32timezone
from keyring.backends import Windows
import msvcrt
keyring.set_keyring(Windows.WinVaultKeyring())
username = getuser()
username = 'cvieites'
userpassword = keyring.get_password("PDFStamper", username)
if (userpassword is None):
    userpassword = passwordcontrol(username)

# User to input file Excel file
print('\n')
print('Loading default request file "PDFStamper list.xlsx"...')
timeflag = timer(3, 'Wait 3 seconds, or press any key now to change the request file.')
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
    msg = 'File Codes.xlsx is open. Close the excel file and re-start.'
    error_handling(msg)

# Read codes excel file with 3 columns: Pack ID, Document, Copies
erroraccessingproquis = 1
while erroraccessingproquis == 1:
    try:
        codeslist = pd.read_excel(outputfile, usecols='E,G,H')
        codes = pd.DataFrame(codeslist)
        erroraccessingproquis = 0
    except Exception:
        from easygui import ccbox
        title = 'Error Accessing Proquis'
        msg = 'Unable to read the "Pack Codes" list from Proquis. Continue and re-enter BTG Password?'
        response = ccbox(msg, title)
        if response is True:
            passwordcontrol(username)
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
    else:
        # Requested as a document pack code
        wrongcode = 1
        for j in range(len(codes['Pack ID'])):
            if (userlist.loc[i, 'Document']) in (codes.loc[j, 'Pack ID']):
                wrongcode = 0
                workinglist_id.append(codes.loc[j, 'Document'])
                workinglist_copies.append(codes.loc[j, 'Copies'])
                workinglist_batch.append(userlist.loc[i, 'BatchNo'])
                # All Forms to be printed double sided (duplex on)
                if '-F' in (codes.loc[j, 'Document']) and '-BR' not in (codes.loc[j, 'Document']):
                    workinglist_duplex.append('D')
                else:
                    workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
    # Error catching a requested wrong code using flag wrongcode = 1
    if wrongcode == 1:
        msg = 'Document Pack Code ' + (userlist.loc[i, 'Document']) + ' is not valid. Check request list.'
        error_handling(msg)

import numpy as np
downloadlist = np.unique(workinglist_id)
print('Your request contains', len(workinglist_id), 'documents, of which, ')
print(str(len(downloadlist)), 'are unique (no repeats) to be retrieved from Proquis: ')
print('\n')
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

# Download files from SSL Proquis
from tqdm import tqdm
print('\n')
print('DOWNLOADING FILES - ' + str(len(downloadlist)) + ' documents')
url = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO='
for x in tqdm(range(len(downloadlist)), unit='file'):
    urlfile = url + str(downloadlist[x])
    outputfile = temppath + str(downloadlist[x]) + '.PDF'
    tqdm.write('Downloading: ' + urlfile)
    download_file(urlfile, outputfile, username, userpassword)

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
    try:
        Stamper(inputpdf, outputpdf)
    except Exception:
        msg = (workinglist_id[i]) + '  1.- Is it a PDF (not Word)?   2.- Incorrect document ID number?  3.-Locked in Proquis?'
        error_handling(msg)
print('\n')
print("Stamping has completed!")
print('\n')

# Send to the printer
print('PRINTING')
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
    time.sleep(3)

# Goodbye message
print('\n')
print('Documents sent to the printer.')
print('All done. - Script terminated. Have a nice day!')
print('\n')

# Ask user to open folder with resulting stamped files
msg = 'All done. - Have a nice day!'
choices = ["Open Temp folder", "Exit"]
from easygui import buttonbox
reply = buttonbox(msg, choices=choices, title='Operation Completed')
if reply == "Open Temp folder":
    path = os.path.realpath(temppath)
    os.startfile(path)
raise SystemExit
