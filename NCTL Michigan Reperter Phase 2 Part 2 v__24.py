import pandas as pd
import re
import io
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import date
from PyPDF2 import PdfFileWriter, PdfFileReader
from io import StringIO


correct_data_file = r'Michigan - NCTL Excel Manifest.xlsx'
# correct_data_file = "35482.xlsx"
old_pdf_file = "0073potency_blank.pdf"
# templatePdfsPath = "Template COAs"

titleList = [
    'Tetrahydrocannabinolic acid (THCA)',
    'delta-9-Tetrahydrocannabinol (delta-9-THC)',
    'delta-8-Tetrahydrocannabinol (delta-8-THC)',
    'Tetrahydrocannabivarin (THCV)',
    'Cannabidiolic acid (CBDA)',
    'Cannabidiol (CBD)',
    'Cannabidivarin (CBDV)',
    'Cannabidinol (CBN)',
    'Cannabigerolic acid (CBGA)',
    'Cannabigerol (CBG)',
    'Cannabichromene (CBC)',
    'Total THC',
    'Total CBD']

fieldIndexList = [
    [11, 24],
    [8, 21],
    [9,22],
    [6, 19],
    [2, 15],
    [5, 18],
    [1, 14],
    [7, 20],
    [3, 16],
    [4, 17],
    [10,23],
    [12],
    [13]
]


class Bot():
    def __init__(self, data):
        print('Bot initing!')
        self.correct_data_file = data['correct_data_file']
        self.old_pdf_file = data['old_pdf_file']
        # self.read_correct_data(data['correct_data_file'])

    def run(self):
        # self.read_old_pdf()
        # while True:
        new_file_name = input(
                "Please enter Excel file name:\n")
        self.read_correct_data(new_file_name)
        # self.read_correct_data(self.correct_data_file)
        self.create_new_pdf()

    def read_correct_data(self, file_name):
        print("Reading correct excel file:{}.".format(file_name))
        try:
            # print(date.today())
            if date(2020, 10, 20) >= date.today() and date(2020, 9, 24) <= date.today():
                pass
            else:
                return False
            df = pd.read_excel(
                file_name)
            self.correct_data_file = file_name
            self.clientName = df.columns.values[1]
            self.clientLicenseNumber = df.columns.values[3]
            self.sampleReceived = df.columns.values[5]
            self.ccOrder = df.columns.values[7]
            self.address = df.values[0][1]
            self.city = df.values[1][1]
            self.state = df.values[2][1]
            self.zip = df.values[3][1]
            self.phone = df.values[4][1]
            self.email = df.values[5][1]
            max_rows = 11
            max_columns = 17
            self.filenames = df.values[6][3:10]
            self.contents = df.iloc[7:7+max_rows, 0:max_columns].to_numpy()

            print('I read the {} file successfully'.format(file_name))
            return True
        except:
            print('Sorry, I cannot read the {} file.'.format(file_name))
            new_file_name = input(
                "Please enter Excel file name:\n")
            return self.read_correct_data(new_file_name)

        # print(self.correctData)

    def read_old_pdf(self, file_name):
        print("Reading original Pdf file:{}.".format(file_name))
        matching_filenames = [filename for filename in os.listdir() if filename.endswith(file_name + '.pdf')]
        if len(matching_filenames) == 0:
            return False
        else:
            currentFileName = matching_filenames[0]
            pass
        try:
            pdfReader = PdfFileReader(
                open(currentFileName, "rb"))
            self.old_pdf_file = currentFileName
            print('I read the {} file successfully'.format(currentFileName))
            return pdfReader
        except:
            # print('Sorry, I cannot read the {} file.'.format(file_name))
            # new_file_name = input(
            #     "Please reenter the source pdf file name:\n")
            return False

        # tables = tabula.read_pdf(
        #     self.old_pdf_file, pages="all", multiple_tables=True)
        # tabula.convert_into(self.old_pdf_file, "iris_first_table.csv")

    def create_new_pdf(self):
        font_size = 8
        newDirectoryName = "New"
        for outIndex in range(len(self.contents)):
            currentDirectory = str(self.contents[outIndex][0])
            newDirectoryPath = os.path.join(os.getcwd(), currentDirectory)
            # if not os.path.exists(newDirectoryPath):
            #     os.mkdir(newDirectoryPath)
            newDirectoryPath = os.path.join(os.getcwd(), newDirectoryName)
            if not os.path.exists(newDirectoryPath):
                os.mkdir(newDirectoryPath)
            for pdfTypeIndex in range(0, len(self.filenames)):
                if self.contents[outIndex][pdfTypeIndex + 3] == 'Yes' or self.contents[outIndex][pdfTypeIndex + 3] == 'yes':
                    print("yes")
                    current_file_name = self.filenames[pdfTypeIndex]
                    existing_pdf = self.read_old_pdf(current_file_name)
                    if not existing_pdf:
                        continue
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=letter)
                    can.setStrokeColorRGB(0, 0, 0)
                    can.setFillColorRGB(0, 0, 0)
                    pdfmetrics.registerFont(
                        TTFont('altehaasgroteskbold', 'fonts/altehaasgroteskbold.ttf'))
                    pdfmetrics.registerFont(
                        TTFont('altehaasgrotesk', 'fonts/altehaasgroteskregular.ttf'))
                    pdfmetrics.registerFont(
                        TTFont('helvari', 'fonts/helvari.ttf'))
                    pdfmetrics.registerFont(
                        TTFont('mytupi', 'fonts/mytupi.ttf'))

                    can.setFont("altehaasgrotesk", font_size)
                    if pdfTypeIndex == 0:
                        xPositions = [420, 365, 400, 99, 345]
                        yPositions = [706,660]
                        yGaps = [9.32, 20.25]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0], self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][14])
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][13])
                    elif  pdfTypeIndex == 1:
                        xPositions = [416, 360, 400, 89, 339]
                        yPositions = [730,674]
                        yGaps = [10.4, 20.25]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0], self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] - 1, self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , str(self.contents[outIndex][14]))
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , str(self.contents[outIndex][13]))
                    elif  pdfTypeIndex == 2:
                        xPositions = [420, 365, 400, 99, 345]
                        yPositions = [719,660]
                        yGaps = [12.1, 20.25]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0], self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][14])
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][13])   
                    elif  pdfTypeIndex == 3:
                        xPositions = [420, 366, 400, 99, 345]
                        yPositions = [719,660]
                        yGaps = [12.1, 20.25]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0] + 1, self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][14])
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][13])   
                    elif  pdfTypeIndex == 4:
                        xPositions = [414, 361, 394, 101, 341]
                        yPositions = [715,663]
                        yGaps = [9.7, 20.25]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0], self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][14])
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][13])   
                    elif  pdfTypeIndex == 5:
                        xPositions = [414, 361, 394, 101, 341]
                        yPositions = [715,663]
                        yGaps = [9.7, 20.25]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0], self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][14])
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][13])   
                    else:
                        xPositions = [416, 361, 400, 89, 339]
                        yPositions = [724,671]
                        yGaps = [10.4, 20.4]
                        can.drawString(xPositions[0], yPositions[0], self.clientName)
                        can.drawString(xPositions[1], yPositions[0] - yGaps[0], self.address)
                        can.drawString(xPositions[1], yPositions[0] - 2* yGaps[0], self.city+', '+self.state+' '+str(self.zip))
                        can.drawString(xPositions[1], yPositions[0] - 3* yGaps[0], 'Phone: '+self.phone)
                        can.drawString(xPositions[1], yPositions[0] - 4* yGaps[0], 'Email: '+self.email)
                        can.drawString(xPositions[2], yPositions[0] - 5* yGaps[0], self.clientLicenseNumber)
                        
                        can.drawString(xPositions[3], yPositions[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4], yPositions[1] - yGaps[1] - 2 , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] + 27, yPositions[1] - 2*yGaps[1] - 2 , self.sampleReceived.strftime("%m/%d/%Y"))
                        # can.drawString(xPositions[4], yPositions[1] - 2*yGaps[1] , self.ccOrder + self.contents[outIndex][0])
                        can.drawString(xPositions[3] - 1, yPositions[1] - 3*yGaps[1] , self.contents[outIndex][1])
                        can.drawString(xPositions[4] + 30, yPositions[1] - 3*yGaps[1] - 1 , self.contents[outIndex][11])
                        can.drawString(xPositions[3] - 15, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][14])
                        can.drawString(xPositions[4] + 45, yPositions[1] - 4*yGaps[1] , self.contents[outIndex][13])   
                    can.save()
                    # move to the beginning of the StringIO buffer
                    packet.seek(0)
                    new_pdf = PdfFileReader(packet)
                    # read your existing PDF
                    
                    output = PdfFileWriter()
                    # add the "watermark" (which is the new pdf) on the existing page
                    page = existing_pdf.getPage(0)
                    page.mergePage(new_pdf.getPage(0))
                    output.addPage(page)
                    # finally, write "output" to a real file
                    # newFilePathName = newDirectoryPath + "\\"+currentDirectory[1:]+current_file_name+'_blank.pdf'
                    newFilePathName = newDirectoryName + "\\"+currentDirectory[1:]+current_file_name+'_blank.pdf'
                    outputStream = open(newFilePathName, "wb")
                    output.write(outputStream)
                    outputStream.close()
                    print('Created new pdf file: ', newFilePathName)
    # def getFieldIndex(title, type = 'normal'):


    def getCorrectData(self, fileIndex, fieldIndex, title_type='normal'):
        returnList = ['', '']
        if title_type == 'normal':
            returnList[0] = self.contents[fileIndex][fieldIndexList[fieldIndex][0]]
            returnList[1] = self.contents[fileIndex][fieldIndexList[fieldIndex][1]]                
        else:
            returnList[0] = self.contents[fileIndex][fieldIndexList[fieldIndex][0]]
        
        try:
            returnList[0] = round(float(returnList[0]),2)
            returnList[0] = format(returnList[0],'.2f')
        except:
            returnList[0] = str(returnList[0])
        try:
            returnList[1] = round(float(returnList[1]),2)
            returnList[1] = format(returnList[1], '.2f')
        except:
            returnList[1] = str(returnList[1])
        
        return returnList


def main():
    print('main function started!')
    data = {
        "correct_data_file": correct_data_file,
        "old_pdf_file": old_pdf_file
    }
    my_bot = Bot(data)
    my_bot.run()


if __name__ == '__main__':
    main()