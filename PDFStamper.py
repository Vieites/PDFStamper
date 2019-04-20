"""
PDFStamper v1.0 alpha
An automated Python 3.7 script to stamp PDFs with watermark and batch numbers
(c) Carlos Vieites - 2019 - All rights reserved
Permision is given for commercial use in all BTG plc facilities
"""

# Import libraries
import os
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
import time
# (unused, future?) from PyPDF2 import PdfFileReader
# (unused, future?) from reportlab.lib import colors
# (unused, future?) from reportlab.pdfbase.ttfonts import TTFont
# (unused, future? copying files) import shutil

# Welcome message
print('Welcome to PDFStamper v1.0 alpha (testing).')
print('\n')
print('This program will stamp a data integrity watermark and will add')
print('a batch number to the PDF you enter in a spreadsheet.')
print('(c) Carlos Vieites - 2019 - All rights reserved')
print('Permision is given for commercial use in all BTG plc facilities')
print('\n')

# User to input file Excel file
userpath = easygui.fileopenbox()
print('Loading request file:')
print(userpath)
print('\n')

# Read excel file document with 4 columns: Document, Batch Number, Prints, Duplex
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

list = df['Document']

for x in range(len(list)):
    if value not in list[x]:
        list[x] = 'FAR-BR-02.' + str(list[x])
    else:
        continue

df['Document'] = list
print(df)
downloadlist = np.unique(list)

print('\n')
print('Your request contains', len(documentlist['Document']), 'documents, of which, ')
print(len(downloadlist), ' are unique (no repeats) to be retrieved from Proquis: ')
print('\n')
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

# Download files from SSL Proquis using Chromedriver
print('\n')
print('Downloading files from Proquis. Authenticate if required.')
t0 = time.time()
chrome_options = webdriver.ChromeOptions()
preferences = {"download.default_directory": 'C:\pdfstampertemp',
               "directory_upgrade": True,
               "download.prompt_for_download": False,
               "safebrowsing.enabled": True}
chrome_options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome(chrome_options=chrome_options,
                          executable_path=r'chromedriver.exe')

url = 'https://proquis.btgplc.com/viewdocument.aspx?DOCNO='
urllist = downloadlist
for x in range(len(downloadlist)):
    urllist[x] = url + str(downloadlist[x])
    driver.get(urllist[x])
    print('Downloading: ', downloadlist[x]),

# Wait for the download to complete (specially for large PDFs)
x1 = 0
while x1 == 0:
    count = 0
    li = os.listdir(temppath)
    for x1 in li:
        if x1.endswith(".crdownload"):
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

# Decrypt downloaded PDFs
# Copy qpdf to temp directory
print('Preparing PDFs ...')
print('\n')

# Change cwd to decrypt
installdir = os.getcwd()
os.chdir('c://pdfstampertemp/')

# Call qpdf to decrypt pdfs
workinglist = np.unique(list)
downloadlist = np.unique(list)
for x in range(len(downloadlist)):
    downloadlist[x] = downloadlist[x] + '.PDF'
    workinglist[x] = 'D' + downloadlist[x]
    subprocess.run(["qpdf.exe", "--decrypt", downloadlist[x], workinglist[x]])
    # os.remove(temppath + downloadlist[x])   (slows down the script)
print('Ready! Stamping now.')
print ('This may take a while according with the size of the PDF. Please wait...')
print('\n')

# Set the cwd back to the install directory
os.chdir(installdir)

# Order the file list by batch number
df = df.sort_values('BatchNo')

# Prep: declare variables and create first stamp headers
inputpdf = 'D'+df['Document']+'.PDF'
outputpdf = 'SD' + df['Document']+'.PDF'
batchno = df['BatchNo']
copies = df['Copies']
duplex = df['Single/Double']
stamp = ['StampP.pdf', 'StampL.pdf']
# Bach number position in page Portrait and Landscape (x,y)
xp = 400
yp = 780
xl = 640
yl = 530

# Create inital stapms for first entry
batch = str(batchno[1])

# Portrait: create a new PDF with Reportlab, insert the text in set location with specified font
packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=A4)
can.setFont('Times-Bold', 16)
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

# Landscape: create a new PDF with Reportlab, insert the text in set location with specified font
packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=A4)
can.setFont('Times-Bold', 16)
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

# Watermark the first PDF
initialf = temppath + inputpdf[0]
inputfile = open(initialf, 'rb')
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
resultPdfFile = open(str(temppath + '0' + outputpdf[0]), 'wb')
pdfWriter.write(resultPdfFile)
inputfile.close()
resultPdfFile.close()

# Rest of the files - create new stamp only if needed
for i in range(1, len(inputpdf)):
    if batchno[i] != batchno[i-1]:
        # Portrait: create a new PDF with Reportlab, insert the text in set location with specified font
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont('Times-Bold', 16)
        can.drawString(xp, yp, str(batchno[i]))
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

        # Landscape: create a new PDF with Reportlab, insert the text in set location with specified font
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont('Times-Bold', 16)
        can.drawString(xl, yl, str(batchno[i]))
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

        # Stamp them
        initialf = temppath + inputpdf[i]
        inputfile = open(initialf, 'rb')
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
        resultPdfFile = open(str(temppath + str(i) + outputpdf[i]), 'wb')
        pdfWriter.write(resultPdfFile)
        inputfile.close()
        resultPdfFile.close()

print("Stamping has completed.")
print('\n')
print('Script terminated. Have a nice day!')
print('\n')

# Ask user to open folder with resulting
text = input("Press ENTER to open the folder, any other key to exit")
if text == "":
    path = os.path.realpath("c://pdfstampertemp")
    os.startfile(path)
else:
    exit()
