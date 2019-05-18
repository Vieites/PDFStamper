"""
PDFStamper v1.95 beta.
An automated Python 3.7 script to stamp PDFs with a watermark and batch numbers,
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
import win32timezone
from easygui import passwordbox
from keyring.backends import Windows
from requests_ntlm import HttpNtlmAuth
# import wmi
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
    printsettings = '"color,fit,' + setduplex + ',' + str(copies) + 'x" '
    args = " -print-to-default -silent -exit-on-print -print-settings " + printsettings, infile
    result = subprocess.run(["SumatraPDF.exe", args])
    return result


# Monitors the printing queue and returns 1 when the job is present, 0 when not
''' NOT IN USE - PRINTER CONTROL
def printer_queue(fileinqueue):
    c = wmi.WMI()
    busyflag = 0
    for printer in c.Win32_Printer():
        for job in c.Win32_PrintJob(DriverName=printer.DriverName):
            if str(fileinqueue) in str(job.document):
                busyflag = 1
    return busyflag
'''


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
print('Welcome to PDFStamper v1.95 beta.')
print('\n')
print('This program will stamp a data integrity watermark and will add')
print('batch numbers to the PDFs you enter in an Excel spreadsheet.')
print('\n')
print('(c) Carlos Vieites - 2019 - All rights reserved')
print('Permision is granted for commercial use in all BTG plc facilities')

# Handle password to access HTTPS Proquis
keyring.set_keyring(Windows.WinVaultKeyring())
username = getpass.getuser()
userpassword = keyring.get_password("PDFStamper", username)
timeflag = timer(3, 'Wait 3 seconds or press any key to enter Password...')
if (userpassword is None) or (timeflag == 1):
    userpassword = passwordbox('Password to access Proquis:')
    keyring.set_password("PDFStamper", username, userpassword)
userpassword = keyring.get_password("PDFStamper", username)

# User to input file Excel file
print('Enter Excel file with the requested documents: ')
userpath = easygui.fileopenbox()
print(userpath)
print('\n')

# Read excel file with 4 columns: Document, Batch Number, Prints, Duplex
documentlist = pd.read_excel(userpath, usecols='A,B,C,D')
df = pd.DataFrame(documentlist)

#  Format column Document numbers to 3 digits
df['Document'] = df['Document'].apply(lambda x: '{0:0>3}'.format(x))
#  Deactivate lib Pandas 'silly' message
pd.options.mode.chained_assignment = None

# If values for BMRs in the spreadsheet are XXX e.g.012 instead of full
# Convert numeric values to string type FAR-BR-02.XXX
value = 'FAR'
df['Document'] = df['Document'].astype(str)
workinglist = df['Document']
for x in range(len(workinglist)):
    if value not in workinglist[x]:
        workinglist[x] = str('FAR-BR-02.' + (workinglist[x]))
    else:
        continue
df['Document'] = workinglist
print(df)
uniquelist = np.unique(workinglist)

print('\n')
print('Your request contains', len(documentlist['Document']), 'documents, of which, ')
print(len(uniquelist), ' are unique (no repeats) to be retrieved from Proquis: ')
print('\n')
print(uniquelist)
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
print('\n')
print('Downloading files from BTG EDMS Proquis. Please wait...')
print('\n')
url = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO='
urllist = url + uniquelist
outputlist = temppath + uniquelist + '.PDF'
for x in range(len(urllist)):
    urlfile = str(urllist[x])
    outputfile = str(outputlist[x])
    print('Downloading: ' + urlfile)
    download_file(urlfile, outputfile, username, userpassword)

workinglist = np.unique(df['Document']) + '.PDF'

# Decrypt downloaded PDFs
print('\n')
print('Preparing PDFs ...')
print('\n')
# Call qpdf to decrypt pdfs
for x in range(len(workinglist)):
    subprocess.run(["qpdf.exe", "--decrypt", (temppath + workinglist[x]), (temppath + 'D' + workinglist[x])])
    workinglist[x] = 'D' + workinglist[x]
    # os.remove(temppath + downloadlist[x])   (slows down the script)
print('Ready! Stamping now.')
print('This may take a while according with the size of the PDF. Please wait...')
print('\n')

# Order the file list by batch number
# df = df.sort_values('BatchNo')

# Stamp decrypted PDFs with batch numbers
# Prepare the list of documents to the full excel list
outputlist = 'SD' + df['Document'] + '.PDF'
workinglist = 'D' + df['Document'] + '.PDF'
batchno = df['BatchNo']
copies = df['Copies']
duplex = df['Single/Double']

# Prep: declare variables to be used to create stamps form blank watermarked
stamp = ['StampP.pdf', 'StampL.pdf']
font_type = 'Times-Bold'
font_size = 16

# Stamp PDF
looprange = len(workinglist)
for i in range(looprange):
    inputpdf = temppath + workinglist[i]
    outputpdf = temppath + str(i) + outputlist[i]
    workinglist[i] = outputpdf
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
    insertbn(font_type, font_size, xp, yp, xl, yl, str(batchno[i]))
    # Call the function to do the stamping
    Stamper(inputpdf, outputpdf)
    print(outputlist[i], ' - Done!')
    # Set the variable to continue the loop
    # previousdoctype = currentdoctype
print('\n')
print("Stamping has completed.")
print('\n')

# Send to the printer
filetoprint = 'SD' + df['Document'] + '.PDF'
for x in range(looprange):
    outputpdf = workinglist[x]
    if duplex[x] == 'D' or 'd':
        setduplex = 'duplex'
    else:
        setduplex = 'simplex'
    nocopies = int(copies[x])
    printpdf(outputpdf, nocopies, setduplex)
    filetoprint[x] = str(x) + str(filetoprint[x])
    print(filetoprint[x] + ' - Printed.')

# Goodbye message
print('\n')
print('Documents sent to the printer.')
print('\n')
print('All done. - Script terminated. Have a nice day!')
print('\n')
# Ask user to open folder with resulting
text = input("Press ENTER to open the output folder, any other key to exit...")
if text == "":
    path = os.path.realpath(temppath)
    os.startfile(path)
raise SystemExit
