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


# correct_data_file = r'Terpene Workbook - Michigan.xlsx'
old_pdf_file = "Michigan Terpenes Flower COA.pdf"

class Bot():
    def __init__(self, data):
        print('Bot initing!')
        # self.correct_data_file = data['correct_data_file']
        self.old_pdf_file = data['old_pdf_file']
        # self.read_correct_data(data['correct_data_file'])

    def run(self):
        # self.read_old_pdf()
        # while True:
        if not "correct_data_file" in globals():
            new_file_name = input(
                    "Please enter LCMS Excel file name:\n")
        else:
            new_file_name = correct_data_file
        self.read_correct_data(new_file_name)
        self.create_new_pdf()

    def read_correct_data(self, file_name):
        print("Reading correct excel file:{}.".format(file_name))
        try:
            # print(date.today())
            if date(2020, 11, 17) >= date.today() and date(2020, 9, 24) <= date.today():
                pass
            else:
                return False
            """Read first sheet"""
            df = pd.read_excel(file_name, sheet_name="wt% Final Reporting Result")
            start_rows = 5
            end_rows = 53
            max_column = 35
            self.filenames_resultpro = df.iloc[start_rows:end_rows,1].to_numpy()
            self.contents_resultpro = df.iloc[start_rows:end_rows,2:2+max_column].to_numpy()
            
            """Read second sheet"""
            df = pd.read_excel(file_name, sheet_name="ug per g Final Reporting Result")

            self.filenames_resultppm = df.iloc[start_rows:end_rows,1].to_numpy()
            self.contents_resultppm = df.iloc[start_rows:end_rows,2:2+max_column].to_numpy()

            """Read third sheet"""

            df = pd.read_excel(file_name, sheet_name="%wt LOQs")

            self.filenames_loqpro = df.iloc[start_rows+1:end_rows+1,1].to_numpy()
            self.contents_loqpro = df.iloc[start_rows+1:end_rows+1,2:2+max_column].to_numpy()

            print('I read the {} file successfully'.format(file_name))

            return True
        except:
            print('Sorry, I cannot read the {} file.'.format(file_name))
            new_file_name = input(
                "Please enter Excel file name:\n")
            return self.read_correct_data(new_file_name)


    def read_old_pdf(self, file_name):
        print("Reading original Pdf file:{}.".format(file_name))
        # matching_filenames = [filename for filename in os.listdir() if filename.startswith('Updated_'+file_name)]
        # if len(matching_filenames) == 0:
        #     return False
        # else:
        #     currentFileName = matching_filenames[0]
        #     pass
        currentFileName = file_name
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
        for pdf_file_name in self.filenames_resultpro:
            fileIndex += 1
            if pdf_file_name == 0 :
                continue
            current_file_name = str(pdf_file_name)[2:]
            print('create new pdf file', current_file_name)
            
            existing_pdf = self.read_old_pdf(old_pdf_file)
            if not existing_pdf:
                continue
            top_height = 497
            height_gap = 14.9
            font_size = 7
            middle_rows = 18
            xPositions = [165, 215, 261, 430, 477, 524]

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


            resultPro = self.contents_resultpro[fileIndex]
            resultPPM = self.contents_resultppm[fileIndex]
            loqPro = self.contents_loqpro[fileIndex]
            for indexRow in range(len(resultPro)):
                try:
                    value = round(float(resultPro[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(resultPro[indexRow])
                
                if indexRow < middle_rows:
                    xPos = xPositions[1]
                    yPos = top_height - indexRow * height_gap
                else:
                    xPos = xPositions[4]
                    yPos = top_height - (indexRow - middle_rows) * height_gap
                can.drawString(xPos, yPos, value)

                try:
                    value = round(float(resultPPM[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(resultPPM[indexRow])
                
                if indexRow < middle_rows:
                    xPos = xPositions[2]
                else:
                    xPos = xPositions[5]
                can.drawString(xPos, yPos, value)
                
                try:
                    value = round(float(loqPro[indexRow]), 2)
                    value = format(value, '.3f')
                except:
                    value = str(loqPro[indexRow])
                
                if indexRow < middle_rows:
                    xPos = xPositions[0]
                else:
                    xPos = xPositions[3]
                can.drawString(xPos, yPos, value)
            
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
            outputStream = open(current_file_name+self.old_pdf_file, "wb")
            output.write(outputStream)
            outputStream.close()
            print('Created new pdf file: ', current_file_name +self.old_pdf_file)

def main():
    print('main function started!')
    data = {
        "old_pdf_file": old_pdf_file
    }
    my_bot = Bot(data)
    my_bot.run()


if __name__ == '__main__':
    main()