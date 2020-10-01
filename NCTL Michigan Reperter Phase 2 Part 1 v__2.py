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


correct_data_file = r'POT BATCH AN99999.xlsx'
# correct_data_file = "35482.xlsx"
old_pdf_file = "0073potency_blank.pdf"
password = "chimichangasandtacos"

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
        self.read_correct_data(self.correct_data_file)
        self.create_new_pdf()

    def read_correct_data(self, file_name):
        print("Reading correct excel file:{}.".format(file_name))
        try:
            # print(date.today())
            if date(2020, 10, 10) >= date.today() and date(2020, 9, 24) <= date.today():
                pass
            else:
                return False
            df = pd.read_excel(
                file_name)
            self.correct_data_file = file_name
            max_rows = 13
            max_columns = 25
            self.filenames = df.iloc[0:max_rows, 0].to_numpy()
            print(self.filenames)
            self.contents = df.iloc[0:max_rows, 0:max_columns].to_numpy()
            print(self.contents)
            self.columns = list(df.columns.values)
            print(self.columns)
            # sub_df = df.iloc[10:23, [16, 32, 31, 20]]
            # self.correctData = sub_df.to_numpy()
            print('I read the {} file successfully'.format(file_name))
            return True
        except:
            print('Sorry, I cannot read the {} file.'.format(file_name))
            new_file_name = input(
                "Please reenter the correct data excel file name:\n")
            return self.read_correct_data(new_file_name)

        # print(self.correctData)

    def read_old_pdf(self, file_name):
        print("Reading original Pdf file:{}.".format(file_name))
        matching_filenames = [filename for filename in os.listdir() if filename.startswith(file_name)]
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
        for pdf_file_name in self.filenames:
            fileIndex += 1
            current_file_name = str(pdf_file_name)[2:]
            print('create new pdf file', current_file_name)
            
            existing_pdf = self.read_old_pdf(current_file_name)
            if not existing_pdf:
                continue
            top_height = 515
            height_gap = 14.37
            font_size = 8
            xPositions = [165, 369, 402, 450]
            # LIMS_ID_positions = [380, 625, 40, 18]

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            # delete area
            can.setStrokeColorRGB(1, 1, 1)
            can.setFillColorRGB(1, 1, 1)
            # can.rect(10, 330, 600, 220, fill=1)
            # can.rect(LIMS_ID_positions[0], LIMS_ID_positions[1],
            #          LIMS_ID_positions[2], LIMS_ID_positions[3], fill=1)
            # insert image
            # mask = [255, 255, 255, 255, 255, 255]
            # can.drawImage('images/table_image.png', x=268, y=327,
            #               width=185, height=180, mask=mask)

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

            index = -1
            for title in titleList:
                index += 1
                height = top_height - round(height_gap * (index + 2))
                can.setStrokeColorRGB(0.8, 0.8, 0.8)
                # can.line(xPositions[0], height, 482, height)
                can.setFont("mytupi", font_size)
                
                if titleList.index(title) < len(titleList) - 2:
                    correctValue = self.getCorrectData(fileIndex, index)
                    # can.drawString(xPositions[0]+1, height+3, title)
                    # can.setFillColorRGB(0.8, 0.8, 0.8)
                    # can.drawString(xPositions[1], height+3, correctValue[0])
                    can.setFillColorRGB(0, 0, 0)
                    can.drawString(xPositions[2], height+3, correctValue[0])
                    can.drawString(xPositions[3], height+3, correctValue[1])
                else:
                    can.setFont("altehaasgroteskbold", font_size)
                    correctValue = self.getCorrectData(fileIndex, index, 'total')
                    # can.drawString(xPositions[0]+1, height+3, title)
                    can.setFillColorRGB(0, 0, 0)
                    can.drawString(xPositions[2], height+3, correctValue[0])
                    can.drawString(xPositions[3], height+4, correctValue[1])

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
            outputStream = open("Updated_"+self.old_pdf_file, "wb")
            output.write(outputStream)
            outputStream.close()
            print('Created new pdf file: ', "Updated_"+self.old_pdf_file)
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