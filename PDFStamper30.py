"""
PDFStamper v2.0 beta.
An automated Python 3.7 script to stamp PDFs with a watermark and batch numbers
(c) Carlos Vieites - 2019 - All rights reserved.
"""
# Import libraries
import os
import time
import easygui
import numpy as np
import pandas as pd
import subprocess
import io
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import requests
from pathlib import Path
import getpass
import keyring
from easygui import passwordbox
from keyring.backends import Windows
from requests_ntlm import HttpNtlmAuth
import msvcrt


# Define functions --------------------------------------------------------------------------------------------
# Download a file from https with authentication handled by getpass and keyring
def download_file(url, outputfile, username, password):
    response = requests.get(url, allow_redirects=True, auth=HttpNtlmAuth(username, password))
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
# End of the define functions section ----------------------------------------------------------------------------


# Welcome message
os.system('color 8f')
print('*****************************************************************************')
print('Welcome to PDFStamper v2.0 beta.')
print('This program will stamp a Data Integrity watermark and will add')
print('batch numbers to the PDFs you enter in an Excel spreadsheet.')
print('(c) Carlos Vieites - 2019 - All rights reserved')
print('Permision is granted for commercial use in all BTG plc facilities')
print('*****************************************************************************')
print('\n')

# Handle password to access HTTPS Proquis
os.system('color 8e')
keyring.set_keyring(Windows.WinVaultKeyring())
username = getpass.getuser()
userpassword = keyring.get_password("PDFStamper", username)
timeflag = timer(3, 'Wait 3 seconds or press any key to enter a new Password...')
if (userpassword is None) or (timeflag == 1):
    userpassword = passwordbox('Password to access Proquis:')
    keyring.set_password("PDFStamper", username, userpassword)
userpassword = keyring.get_password("PDFStamper", username)

# User to input file Excel file
print('Loading default request file "PDFStamper list.xlsx".')
timeflag = timer(3, 'Wait 3 seconds or press any key to change the request file...')
if timeflag == 1:
    userpath = easygui.fileopenbox()
    print(userpath)
else:
    userpath = 'PDFStamper list.xlsx'

# Read user excel file with 4 columns: Document, Batch Number, Prints, Duplex
inputlist = pd.read_excel(userpath, usecols='A,B,C,D')
df = pd.DataFrame(inputlist)
userlist = pd.DataFrame(inputlist)

# Download document pack codes list from Proquis FAR-SP-02.001-F11
codesurl = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO=FAR-SP-02.001-F11'
outputfile = 'Codes.xlsx'
download_file(codesurl, outputfile, username, userpassword)

# Read codes excel file with 3 columns: Pack ID, Document, Copies
codeslist = pd.read_excel(outputfile, usecols='E,G,H')
codes = pd.DataFrame(codeslist)

# Construct the working lists to download and PDF stamp using the codes list
os.system('color 8f')
workinglist_id = []
workinglist_copies = []
workinglist_batch = []
workinglist_duplex = []

for i in range(len(userlist['Document'])):
    if 'FAR-' in (userlist.loc[i, 'Document']):
        workinglist_id.append(userlist.loc[i, 'Document'])
        workinglist_copies.append(userlist.loc[i, 'Copies'])
        workinglist_batch.append(userlist.loc[i, 'BatchNo'])
        workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
    else:
        for j in range(len(codes['Pack ID'])):
            if (userlist.loc[i, 'Document']) in (codes.loc[j, 'Pack ID']):
                workinglist_id.append(codes.loc[j, 'Document'])
                workinglist_copies.append(codes.loc[j, 'Copies'])
                workinglist_batch.append(userlist.loc[i, 'BatchNo'])
                workinglist_duplex.append(userlist.loc[i, 'Single/Double'])
downloadlist = np.unique(workinglist_id)
print('\n')
print('Your request contains', len(workinglist_id), 'documents, of which, ')
print(len(downloadlist), ' are unique (no repeats) to be retrieved from Proquis: ')
print(downloadlist)
print('\n')

# Prep a temp folder for downloads
# Check if the temp folder exists otherwise, create it
temppath = 'C:/pdfstampertemp/'
if not os.path.exists(temppath):
    os.makedirs(temppath)
# Check temp folder is empty otherwise, clean it
if os.listdir(temppath) != []:
    fileList = os.listdir(temppath)
    for fileName in fileList:
        os.remove(temppath + fileName)

# Download files from SSL Proquis
os.system('color 8a')
print('\n')
print('Downloading files from BTG EDMS Proquis. Please wait...')
url = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO='
for x in range(len(downloadlist)):
    urlfile = url + str(downloadlist[x])
    outputfile = temppath + str(downloadlist[x]) + '.PDF'
    print('Downloading: ' + urlfile)
    download_file(urlfile, outputfile, username, userpassword)

# Decrypt downloaded PDFs
print('\n')
print('Preparing PDFs ...',)
# Call qpdf to decrypt pdfs
for x in range(len(downloadlist)):
    print('Preparing ' + downloadlist[x])
    inputfile = temppath + downloadlist[x] + '.PDF'
    outputfile = temppath + 'D' + downloadlist[x] + '.PDF'
    subprocess.run(["qpdf.exe", "--decrypt", inputfile, outputfile])
    print(' - Done!')

# Stamp decrypted PDFs with batch numbers
print('\n')
print('Ready! Stamping now.')
print('This may take a while according with the size of the PDF. Please wait...')
# Prep: declare variables to be used to create stamps form blank watermarked
stamp = ['StampP.pdf', 'StampL.pdf']
font_type = 'Times-Bold'
font_size = 16

# Stamp PDF
looprange = len(workinglist_id)
for i in range(looprange):
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
    Stamper(inputpdf, outputpdf)
    print(workinglist_id[i], ' - Done!')

print('\n')
print("Stamping has completed!")
print('\n')

# Send to the printer
print('Printing documents.')
for x in range(looprange):
    if workinglist_duplex[x] == 'D' or workinglist_duplex[x] == 'd':
        setduplex = 'duplex'
    else:
        setduplex = 'simplex'
    filetoprint = temppath + str(x) + 'SD' + str(workinglist_id[x]) + '.PDF'
    nocopies = int(workinglist_copies[x])
    print('Printing ' + filetoprint + '...')
    printpdf(filetoprint, nocopies, setduplex)
    print(filetoprint + ' - Printed.')
    # Give some time for the file to reach the printing buffer
    time.sleep(2)

# Goodbye message
os.system('color 8f')
print('\n')
print('Documents sent to the printer.')
print('\n')
print('All done. - Script terminated. Have a nice day!')
print('\n')
# Ask user to open folder with resulting stamped files
text = input("Press ENTER to open the output folder, any other key to exit...")
if text == "":
    path = os.path.realpath(temppath)
    os.startfile(path)
raise SystemExit
