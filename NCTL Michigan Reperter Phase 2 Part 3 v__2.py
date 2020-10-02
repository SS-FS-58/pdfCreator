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


correct_data_file = r'Pesticide Workbook - LCMS.xlsx'
correct_data_file1 = r'Pesticide Workbook - GCMS.xlsx'

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
                "Please enter LCMS Excel file name:\n")
        self.read_correct_data(new_file_name)
        # self.read_correct_data(correct_data_file)
        new_file_name = input(
                "Please enter GCMS Excel file name:\n")
        self.read_correct_data1(correct_data_file1)
        self.create_new_pdf()

    def read_correct_data(self, file_name):
        print("Reading correct excel file:{}.".format(file_name))
        try:
            # print(date.today())
            if date(2020, 10, 30) >= date.today() and date(2020, 9, 24) <= date.today():
                pass
            else:
                return False
            df = pd.read_excel(file_name, sheet_name="ug per g LOQ's")

            max_rows = 48
            max_column = 53
            self.filenames_ug = df.iloc[7:max_rows+7,1].to_numpy()
            self.contents_ug = df.iloc[7:max_rows+7,2:2+max_column].to_numpy()

            df = pd.read_excel(file_name, sheet_name="Final Reporting Results")

            self.filenames_final = df.iloc[7:max_rows+7,1].to_numpy()
            self.contents_final = df.iloc[7:max_rows+7,2:2+max_column].to_numpy()

            print('I read the {} file successfully'.format(file_name))
            return True
        except:
            print('Sorry, I cannot read the {} file.'.format(file_name))
            new_file_name = input(
                "Please enter Excel file name:\n")
            return self.read_correct_data(new_file_name)

        # print(self.correctData)
    def read_correct_data1(self, file_name):
        print("Reading correct excel file:{}.".format(file_name))
        try:
            # print(date.today())
            if date(2020, 10, 30) >= date.today() and date(2020, 9, 24) <= date.today():
                pass
            else:
                return False
            df = pd.read_excel(file_name, sheet_name="ug per g LOQ's")

            max_rows = 48
            max_column = 5
            self.filenames_ug1 = df.iloc[7:max_rows+7,1].to_numpy()
            self.contents_ug1 = df.iloc[7:max_rows+7,2:2+max_column].to_numpy()

            df = pd.read_excel(file_name, sheet_name="Final Reporting Results")

            self.filenames_final1 = df.iloc[7:max_rows+7,1].to_numpy()
            self.contents_final1= df.iloc[7:max_rows+7,2:2+max_column].to_numpy()

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
        matching_filenames = [filename for filename in os.listdir() if filename.startswith('Updated_'+file_name)]
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
        fileIndex = -1
        for pdf_file_name in self.filenames_ug:
            fileIndex += 1
            if pdf_file_name == 0 :
                continue
            current_file_name = str(pdf_file_name)[2:]
            print('create new pdf file', current_file_name)
            
            existing_pdf = self.read_old_pdf(current_file_name)
            if not existing_pdf:
                continue
            top_height = 513
            height_gap = 8.55
            font_size = 7
            middle_rows = 29
            xPositions = [213, 124, 494, 405]

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            
            can.setStrokeColorRGB(0, 0, 0)
            can.setFillColorRGB(0, 0, 0)

            # Title
            pdfmetrics.registerFont(
                TTFont('altehaasgroteskbold', 'fonts/altehaasgroteskbold.ttf'))
            pdfmetrics.registerFont(
                TTFont('altehaasgrotesk', 'fonts/altehaasgroteskregular.ttf'))
            pdfmetrics.registerFont(
                TTFont('helvari', 'fonts/helvari.ttf'))
            pdfmetrics.registerFont(
                TTFont('mytupi', 'fonts/mytupi.ttf'))

            can.setFont("altehaasgrotesk", font_size)


            loqPPM = self.contents_ug[fileIndex]
            resultPPM = self.contents_final[fileIndex]
            for indexRow in range(len(loqPPM)):
                try:
                    value = round(float(loqPPM[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(loqPPM[indexRow])
                
                if indexRow < middle_rows:
                    xPos = xPositions[0]
                    yPos = top_height - indexRow * height_gap
                else:
                    xPos = xPositions[2]
                    yPos = top_height - (indexRow - middle_rows) * height_gap
                can.drawString(xPos, yPos, value)

                try:
                    value = round(float(resultPPM[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(resultPPM[indexRow])
                
                if indexRow < middle_rows:
                    xPos = xPositions[1]
                else:
                    xPos = xPositions[3]
                can.drawString(xPos, yPos, value)

            loqPPM1 = self.contents_ug1[fileIndex]  
            resultPPM1 = self.contents_final1[fileIndex] 
            for indexRow in range(len(loqPPM1)):
                try:
                    value = round(float(loqPPM1[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(loqPPM1[indexRow])
                
                xPos = xPositions[2]
                yPos -= height_gap
                can.drawString(xPos, yPos, value)
                try:
                    value = round(float(resultPPM1[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(resultPPM1[indexRow])
                can.drawString(xPositions[3], yPos, value)
            # can.drawString(xPositions[3], height+3, "correctValue[1]")

            # can.drawString(LIMS_ID_positions[0]+7,
            #                LIMS_ID_positions[1]+6, self.LIMS_ID)
            # can.line(xPositions[0], height, 482, height)
            # can.setFont("altehaasgroteskbold", 8)
            # can.drawString(xPositions[0]+1, height+3, "Analyte")
            # can.setFillColorRGB(0.8, 0.8, 0.8)
            # can.drawString(xPositions[1], height+3, "LOQ")
            # can.setFillColorRGB(0, 0, 0)
            # can.drawString(xPositions[2], height+3, "Result")
            # can.drawString(xPositions[3], height+3, "Result")

            # index = -1
            # for title in titleList:
            #     index += 1
            #     height = top_height - round(height_gap * (index + 2))
            #     can.setStrokeColorRGB(0.8, 0.8, 0.8)
            #     # can.line(xPositions[0], height, 482, height)
            #     can.setFont("mytupi", font_size)
                
            #     if titleList.index(title) < len(titleList) - 2:
            #         correctValue = self.getCorrectData(fileIndex, index)
            #         # can.drawString(xPositions[0]+1, height+3, title)
            #         # can.setFillColorRGB(0.8, 0.8, 0.8)
            #         # can.drawString(xPositions[1], height+3, correctValue[0])
            #         can.setFillColorRGB(0, 0, 0)
            #         can.drawString(xPositions[2], height+3, correctValue[0])
            #         can.drawString(xPositions[3], height+3, correctValue[1])
            #     else:
            #         can.setFont("altehaasgroteskbold", font_size)
            #         correctValue = self.getCorrectData(fileIndex, index, 'total')
            #         # can.drawString(xPositions[0]+1, height+3, title)
            #         can.setFillColorRGB(0, 0, 0)
            #         can.drawString(xPositions[2], height+3, correctValue[0])
            #         can.drawString(xPositions[3], height+4, correctValue[1])

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
            outputStream = open("Re_"+self.old_pdf_file, "wb")
            output.write(outputStream)
            outputStream.close()
            print('Created new pdf file: ', "Re_"+self.old_pdf_file)

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