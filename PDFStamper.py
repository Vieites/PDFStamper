"""
PDFStamper v1.3 alpha
An automated Python 3.7 script to stamp PDFs with a watermark and batch numbers
(c) Carlos Vieites - 2019 - All rights reserved
"""
# Import libraries
import os
import time
import datetime
import easygui
import numpy as np
import pandas as pd
from selenium import webdriver
import subprocess
import io
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import win32api
import win32print
from selenium import webdriver
# (unused, future?) from reportlab.lib import colors
# (unused, future?) from reportlab.pdfbase.ttfonts import TTFont
# (unused, future? copying files) import shutil

# Welcome message
print('Welcome to PDFStamper v1.3 alpha (testing).')
print('\n')
print('This program will stamp a data integrity watermark and will add')
print('the batch number to the PDF you enter from an Excel spreadsheet.')
print('\n')
print('(c) Carlos Vieites - 2019 - All rights reserved')
print('Permision is granted for commercial use in all BTG plc facilities')
print('\n')

# User to input file Excel file
print('Enter request MS Excel file:')
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
workinglist = np.unique(workinglist)

print('\n')
print('Your request contains', len(documentlist['Document']), 'documents, of which, ')
print(len(workinglist), ' are unique (no repeats) to be retrieved from Proquis: ')
print('\n')
print(workinglist)
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

# Download files from SSL Proquis using Chromedriver
print('\n')
print('Downloading files from Proquis. Authenticate if required.')
t0 = time.time()
chrome_options = webdriver.ChromeOptions()
options = chrome_options
preferences = {"download.default_directory": 'C:\pdfstampertemp',
               "directory_upgrade": True,
               "download.prompt_for_download": False,
               "safebrowsing.enabled": True}
chrome_options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome(options=options,
                          executable_path=r'chromedriver.exe')

url = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO='
urllist = workinglist
for x in range(len(urllist)):
    urllist[x] = str(url + workinglist[x])
    driver.get(urllist[x])
    print('Downloading: ', workinglist[x]),
    time.sleep(0.5)
    # Chromedriver does not work if time is not given to process the 1st job
    if x == 0:
        time.sleep(3)
# Wait for the download to complete (specially for large PDFs)
x1 = 0
while x1 == 0:
    count = 0
    li = os.listdir(temppath)
    for x1 in li:
        if x1.endswith(".crdownload") or x1.endswith(".tmp"):
            count = count + 1
            print('{} seconds'.format(int(time.time() - t0)), end='\r')
    if count == 0:
        x1 = 1
    else:
        x1 = 0

print('\n')
elapsedtime = int(time.time() - t0)
print('Download completed in', str(datetime.timedelta(seconds=elapsedtime)))
print('\n')
driver.quit()
workinglist = np.unique(df['Document']) + '.PDF'

# Decrypt downloaded PDFs
print('Preparing PDFs ...')
print('\n')

# Decrypt files
# Call qpdf to decrypt pdfs
for x in range(len(workinglist)):
    subprocess.run(["qpdf.exe", "--decrypt", (temppath + workinglist[x]), (temppath + 'D' + workinglist[x])])
    workinglist[x] = 'D' + workinglist[x]
    # os.remove(temppath + downloadlist[x])   (slows down the script)
print('Ready! Stamping now.')
print('This may take a while according with the size of the PDF. Please wait...')
print('\n')

# Order the file list by batch number
df = df.sort_values('BatchNo')


# Inject batch number in already watermarked blank Portrait and Landscape PDFs
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
    output.write(outputStream)
    outputStream.close()
    return


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
    return


# Print function with Duplex selection
def printpdf(infile, setduplex):
    name = win32print.GetDefaultPrinter()
    # printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_ADMINISTER}
    printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}
    handle = win32print.OpenPrinter(name, printdefaults)
    level = 2
    attributes = win32print.GetPrinter(handle, level)
    print("Old Duplex = %d" % attributes['pDevMode'].Duplex)

    # attributes['pDevMode'].Duplex = 1    no flip (single sided)
    # attributes['pDevMode'].Duplex = 2       flip up
    # attributes['pDevMode'].Duplex = 3       flip over (doble sided)
    attributes['pDevMode'].Duplex = setduplex
    # 'SetPrinter' fails because of 'Access is denied.'
    # But the attribute 'Duplex' is set correctly
    try:
        win32print.SetPrinter(handle, level, attributes, 0)
    except:
        # THIS PRINT LINE DOWN HERE IS NOT NEEDED?
        print('win32print.SetPrinter: set Duplex')
    res = win32api.ShellExecute(0, 'print', infile, None, '.', 0)
    win32print.ClosePrinter(handle)
    return


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

# Bach number position in page Portrait and Landscape (x,y)
xp = 420
yp = 780
xl = 640
yl = 530

# Stamp first PDF
inputpdf = temppath + workinglist[0]
outputpdf = temppath + '0' + outputlist[0]
workinglist[0] = outputpdf
insertbn(font_type, font_size, xp, yp, xl, yl, str(batchno[0]))
Stamper(inputpdf, outputpdf)
print(outputlist[0], ' - Done!')

# Stamp the rest of the PDFs
looprange = (len(workinglist)) - 1
for x in range(looprange):
    inputpdf = temppath + workinglist[x + 1]
    outputpdf = temppath + str(x+1) + outputlist[x + 1]
    workinglist[x+1] = outputpdf
    # Only create new batchno stamp for new numbers, otherwise use existing
    if (batchno[x+1] != batchno[x]):
        insertbn(font_type, font_size, xp, yp, xl, yl, str(batchno[x + 1]))
    Stamper(inputpdf, outputpdf)
    print(outputlist[x + 1], ' - Done!')

print('\n')
print("Stamping has completed.")
print('\n')

# Send to the printer
looprange = looprange + 1
for x in range(looprange):
    print('Printing ', workinglist[x])
    outputpdf = workinglist[x]
    if duplex[x] == 'S' or 's':
        setduplex = 1
    else:
        setduplex = 3
    nocopies = int(copies[x])
    for k in range(nocopies):
        printpdf(outputpdf, setduplex)

# Goodbye message
print('\n')
print('Documents sent to the printer.')
print('\n')
print('All done. - Script terminated. Have a nice day!')
print('\n')
# Ask user to open folder with resulting
text = input("Press ENTER to open the folder, any other key to exit")
if text == "":
    path = os.path.realpath(temppath)
    os.startfile(path)
else:
    exit()
