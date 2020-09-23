import pandas as pd
import re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import HTMLConverter
# from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# import tabula

correct_data_file = "35482.xlsx"
old_pdf_file = "NCAL Draft Plant Material CoA.pdf"
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
        self.correct_data_file = data['correct_data_file']
        self.old_pdf_file = data['old_pdf_file']
        self.read_correct_data()
        self.read_old_pdf()
        self.create_new_pdf()

    def read_correct_data(self):
        print("read excel file")
        df = pd.read_excel(
            self.correct_data_file)
        sub_df = df.iloc[10:23, [16, 32, 31, 20]]
        self.correctData = sub_df.to_numpy()
        print(self.correctData)

    def read_old_pdf(self):
        print('read old pdf file')
        # tables = tabula.read_pdf(
        #     self.old_pdf_file, pages="all", multiple_tables=True)
        # tabula.convert_into(self.old_pdf_file, "iris_first_table.csv")

    def create_new_pdf(self):
        print('create new pdf file')
        height = 504
        height_gap = 12
        xPositions = [165, 369, 413, 458]

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        # delete area
        can.setStrokeColorRGB(1, 1, 1)
        can.setFillColorRGB(1, 1, 1)
        can.rect(10, 330, 600, 220, fill=1)
        # insert image
        mask = [255, 255, 255, 255, 255, 255]
        can.drawImage('images/table_image.png', x=268, y=327,
                      width=185, height=180, mask=mask)

        can.setStrokeColorRGB(0, 0, 0)
        can.setFillColorRGB(0, 0, 0)

        # Title
        pdfmetrics.registerFont(
            TTFont('altehaasgroteskbold', 'fonts/altehaasgroteskbold.ttf'))
        pdfmetrics.registerFont(
            TTFont('altehaasgrotesk', 'fonts/altehaasgroteskregular.ttf'))

        can.line(xPositions[0], height, 482, height)
        can.setFont("altehaasgroteskbold", 8)
        can.drawString(xPositions[0]+1, height+3, "Analyte")
        can.setFillColorRGB(0.8, 0.8, 0.8)
        can.drawString(xPositions[1], height+3, "LOQ")
        can.setFillColorRGB(0, 0, 0)
        can.drawString(xPositions[2], height+3, "Result")
        can.drawString(xPositions[3], height+3, "Result")

        for title in titleList:
            height -= height_gap
            can.setStrokeColorRGB(0.8, 0.8, 0.8)
            can.line(xPositions[0], height, 482, height)
            can.setFont("altehaasgrotesk", 8)
            if titleList.index(title) == 0:
                can.setFillColorRGB(0.8, 0.8, 0.8)
                can.drawString(xPositions[1], height+3, "%")
                can.setFillColorRGB(0, 0, 0)
                can.drawString(xPositions[2], height+3, "%")
                can.drawString(xPositions[3], height+3, "mg/g")
            elif titleList.index(title) < len(titleList) - 2:
                correctValue = self.getCorrectData(title)
                can.drawString(xPositions[0]+1, height+3, title)
                can.setFillColorRGB(0.8, 0.8, 0.8)
                can.drawString(xPositions[1], height+3, correctValue[0])
                can.setFillColorRGB(0, 0, 0)
                can.drawString(xPositions[2], height+3, correctValue[1])
                can.drawString(xPositions[3], height+3, correctValue[2])
            else:
                can.setFont("altehaasgroteskbold", 8)
                correctValue = self.getCorrectData(title, 'total')
                can.drawString(xPositions[0]+1, height+3, title)
                can.setFillColorRGB(0, 0, 0)
                can.drawString(xPositions[2], height+3, correctValue[1])
                can.drawString(xPositions[3], height+3, correctValue[2])

        can.save()

        # move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        # read your existing PDF
        existing_pdf = PdfFileReader(
            open(old_pdf_file, "rb"))
        output = PdfFileWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        # finally, write "output" to a real file
        outputStream = open("destination.pdf", "wb")
        output.write(outputStream)
        outputStream.close()

    def getCorrectData(self, title, title_type='normal'):
        returnList = ['', '', '']
        for correct_data in self.correctData:
            print(correct_data)
            keyValue = re.sub(r"\s+$", "", correct_data[0])[-5:]
            if title_type == 'normal':
                keyValue += ')'
            if keyValue in title:
                returnList[0] = str(round(float(correct_data[1]), 4))
                if returnList[0] == '0.0':
                    returnList[0] = 'ND'
                returnList[1] = str(round(float(correct_data[2]), 3))
                if returnList[1] == '0.0':
                    returnList[1] = 'ND'
                returnList[2] = str(round(float(correct_data[3]), 3))
                if returnList[2] == '0.0':
                    returnList[2] = 'ND'

        return returnList


def main():
    print('main function started!')
    data = {
        "correct_data_file": correct_data_file,
        "old_pdf_file": old_pdf_file
    }
    my_bot = Bot(data)


if __name__ == '__main__':
    main()
