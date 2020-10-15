import pandas as pd
import re
# from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
# from pdfminer.converter import HTMLConverter
# # from pdfminer.converter import TextConverter
# from pdfminer.layout import LAParams
# from pdfminer.pdfpage import PDFPage
from io import StringIO

from PyPDF2 import PdfFileWriter, PdfFileReader
import io
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import date
# import tabula

correct_data_file = "36962.xlsx"
old_pdf_file = "36962.pdf"
password = "chimichangasandtacos"
titleList = [
    '',
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
            if date(2020, 10, 31) >= date.today() and date(2020, 9, 24) <= date.today():
                pass
            else:
                return False
            df = pd.read_excel(
                file_name, sheet_name="Product")
            self.correct_data_file = file_name
            self.density = str(float(df.values[5][5]))
            print('density :', self.density)
            sub_df = df.iloc[10:23, [10, 11, 12, 13]]
            self.correctData = sub_df.to_numpy()
            # self.compund = df.iloc[10:23, 10].to_numpy()
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
        try:
            pdfReader = PdfFileReader(
                open(file_name, "rb"))
            if pdfReader.isEncrypted:
                try:
                    pdfReader.decrypt(password)
                except NotImplementedError:
                    command = f"qpdf --password='{password}' --decrypt {file_name} {file_name+'_decrypt'};"
                    os.system(command)
                    with open(file_name+'_decrypt', mode='rb') as fp:
                        pdfReader = PdfFileReader(fp)
            self.old_pdf_file = file_name
            print('I read the {} file successfully'.format(file_name))
            return pdfReader
        except:
            print('Sorry, I cannot read the {} file.'.format(file_name))
            new_file_name = input(
                "Please reenter the source pdf file name:\n")
            return self.read_old_pdf(new_file_name)

        # tables = tabula.read_pdf(
        #     self.old_pdf_file, pages="all", multiple_tables=True)
        # tabula.convert_into(self.old_pdf_file, "iris_first_table.csv")

    def create_new_pdf(self):
        print('create new pdf file')
        height = 507
        height_gap = 12
        opacity_font = 0.6
        xPositions = [243, 439, 500, 541]
        density_positions = [240, 270, 400, 18]

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        # delete area
        can.setStrokeColorRGB(1, 1, 1)
        can.setFillColorRGB(1, 1, 1)
        can.rect(230, 330, 600, 220, fill=1)
        can.rect(density_positions[0], density_positions[1],
                 density_positions[2], density_positions[3], fill=1)
        # insert image
        mask = [255, 255, 255, 255, 255, 255]
        can.drawImage('images/table_image.png', x=345, y=330,
                      width=175, height=170, mask=mask)

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

        can.setFont("altehaasgrotesk", 10)
        can.setFillColorRGB(opacity_font, opacity_font, opacity_font)
        can.drawString(density_positions[0]+3,
                       density_positions[1]+6, 'Results reported in mg/1 fl oz are based on a density of '+ self.density+'.')
        can.setFont("altehaasgroteskbold", 7)
        can.drawString(xPositions[1], height, "LOQ")
        can.setFillColorRGB(0, 0, 0)
        can.line(xPositions[0], height-3, xPositions[3]+25, height-3)
        # can.setFont("altehaasgroteskbold", 8)
        can.drawString(xPositions[0]+1, height, "Analyte")
        can.drawString(xPositions[2] - 5, height, "Result")
        can.drawString(xPositions[3], height, "Result")

        for title in titleList:
            height -= height_gap
            can.setStrokeColorRGB(opacity_font, opacity_font, opacity_font)
            can.line(xPositions[0], height - 3, xPositions[3]+25, height -3)
            can.setFont("mytupi", 7)
            if titleList.index(title) == 0:
                can.setFillColorRGB(opacity_font, opacity_font, opacity_font)
                can.drawString(xPositions[1], height, "    %")
                can.setFillColorRGB(0, 0, 0)
                can.drawString(xPositions[2], height, "  %")
                can.drawString(xPositions[3], height, "mg/g")
            elif titleList.index(title) < len(titleList) - 2:
                correctValue = self.getCorrectData(title)
                can.drawString(xPositions[0]+1, height, title)
                can.setFillColorRGB(opacity_font, opacity_font, opacity_font)
                can.drawString(xPositions[1], height, correctValue[0])
                can.setFillColorRGB(0, 0, 0)
                can.drawString(xPositions[2], height, correctValue[1])
                can.drawString(xPositions[3], height, correctValue[2])
            else:
                can.setFont("altehaasgroteskbold", 7)
                correctValue = self.getCorrectData(title, 'total')
                can.drawString(xPositions[0]+1, height, title)
                can.setFillColorRGB(0, 0, 0)
                can.drawString(xPositions[2], height, correctValue[1])
                can.drawString(xPositions[3], height, correctValue[2])

        can.save()

        # move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        # read your existing PDF
        existing_pdf = self.read_old_pdf(self.old_pdf_file)
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

    def getCorrectData(self, title, title_type='normal'):
        returnList = ['', '', '']
        for correct_data in self.correctData:
            # print(correct_data)
            keyValue = re.sub(r"\s+$", "", correct_data[0])[-6:]
            if keyValue == ' 9THC)':
                keyValue = '9-THC)'
            if keyValue == ' 8THC)':
                keyValue = '8-THC)'
            if keyValue == 'HCA-A)':
                keyValue = '(THCA)'
            if keyValue in title:
                try:
                    returnList[0] = str(round(float(correct_data[1]), 3))
                    if returnList[0] == '0.0':
                        returnList[0] = 'ND'
                except:
                    returnList[0] = 'ND'
                try:
                    returnList[1] = str(round(float(correct_data[2]), 2))
                    if returnList[1] == '0.0':
                        returnList[1] = 'ND'
                except:
                    returnList[1] = 'ND'
                try:
                    returnList[2] = str(round(float(correct_data[3]), 2))
                    if returnList[2] == '0.0':
                        returnList[2] = 'ND'
                except:
                    returnList[2] = 'ND'
                break

        return returnList


def main():
    print('main function started!')
    data = {
        "old_pdf_file": old_pdf_file
    }
    my_bot = Bot(data)
    my_bot.run()


if __name__ == '__main__':
    main()
