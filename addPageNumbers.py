from PyPDF2 import  PdfFileReader, PdfFileWriter
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def createPageWithNumbers (pageNum):
    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.drawString(260, 10, "Page "+str(pageNum))
    can.save()
    return packet


# read your existing PDF
existing_pdf = PdfFileReader(open("C:/MergePDF/Output/ConsolidatedGroupPack.pdf", 'rb'))
output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
for pagenum in range(existing_pdf.numPages):
    pageObj = existing_pdf.getPage(pagenum)
    #pdfWriter.addPage(pageObj)
    page = existing_pdf.getPage(pagenum)
    packet=createPageWithNumbers(pagenum+1)
    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)

# finally, write "output" to a real file
outputStream = open("C:/MergePDF/Output/destination.pdf", "wb")
output.write(outputStream)
outputStream.close()

# actual piece of code
# packet = io.BytesIO()
# # create a new PDF with Reportlab
# can = canvas.Canvas(packet, pagesize=letter)
# can.drawString(260, 10, '1')
# can.save()
#
# #move to the beginning of the StringIO buffer
# packet.seek(0)
# new_pdf = PdfFileReader(packet)
# # read your existing PDF
# existing_pdf = PdfFileReader(open("C:/MergePDF/Output/ConsolidatedGroupPack.pdf", 'rb'))
# output = PdfFileWriter()
# # add the "watermark" (which is the new pdf) on the existing page
# page = existing_pdf.getPage(0)
# page.mergePage(new_pdf.getPage(0))
# output.addPage(page)
# # finally, write "output" to a real file
# outputStream = open("C:/MergePDF/Output/destination.pdf", "wb")
# output.write(outputStream)
# outputStream.close()