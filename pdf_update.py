from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
# from reportlab.lib.units import inch

titleList = [
    'Tetrahydrocannabinolic acid (THCA)',
    'delta-9-Tetrahydrocannabinol (delta-9-THC)',
    'delta-8-Tetrahydrocannabinol (delta-8-THC)',
    'Tetrahydrocannabivarin (THCV)',
    'Cannabidiolic acid (CBDA)',
    'Cannabidiol (CBD)',
    'Cannabidivarin (CBDB)',
    'Cannabidinol (CBN)',
    'Cannabigerolic acid (CBGA)',
    'Cannabigerol (CBG)',
    'Cannabichromene (CBC)',
    'Total THC',
    'Total CBD']
packet = io.BytesIO()
# create a new PDF with Reportlab
can = canvas.Canvas(packet, pagesize=letter)
# can.translate(inch, inch)
# can.drawString(10, 100, "Hello world")
mask = [255, 255, 255, 255, 255, 255]
can.drawImage('table_image.png', x=268, y=327,
              width=185, height=180, mask=mask)


can.setStrokeColorRGB(1, 1, 1)
can.setFillColorRGB(1, 1, 1)
# can.rect(10, 330, 600, 220, fill=1)

can.setStrokeColorRGB(0, 0, 0)
can.setFillColorRGB(0, 0, 0)

can.line(165, 504, 482, 504)

can.setStrokeColorRGB(0.8, 0.8, 0.8)
for height in range(504, 325, -12):
    can.line(165, height, 482, height)

# can.setFillColorRGB(0, 0, 0)
can.setFont("Times-Bold", 8)
can.drawString(166, 507, "Analyse")
can.setFillColorRGB(0.8, 0.8, 0.8)
can.drawString(369, 507, "LOQ")
can.setFillColorRGB(0, 0, 0)
can.drawString(413, 507, "Result")
can.drawString(458, 507, "Result")
can.save()

# move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)
# read your existing PDF
existing_pdf = PdfFileReader(open("NCAL Final Plant Material CoA.pdf", "rb"))
output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.getPage(0)
page.mergePage(new_pdf.getPage(0))
output.addPage(page)
# finally, write "output" to a real file
outputStream = open("destination.pdf", "wb")
output.write(outputStream)
outputStream.close()
