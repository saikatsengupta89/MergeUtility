import PyPDF2
import os
import re
import shutil


def atoi(text):
    return int(text) if text.isdigit() else text

def natural_keys(text):
    return [atoi(c) for c in re.split(r'(\d+)', text)]

def mergePDF(inputLocation, outputLocation, outputFileName):
    # move to the input directory and create a list of all pdfs needs to be merged
    os.chdir(inputLocation)

    pdfs_list = []
    for filename in os.listdir('.'):
        if filename.endswith('.pdf'):
            pdfs_list.append(filename)

    pdfs_list.sort(key=natural_keys)
    #print(pdfs_list)

    pdfWriter = PyPDF2.PdfFileWriter()
    # loop through all the pdfs and merge them one by one
    for filename in pdfs_list:
        # rb for read binary format
        pdfFile = open(filename, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        # opening each page in the pdf
        for pagenum in range(pdfReader.numPages):
            pageObj = pdfReader.getPage(pagenum)
            pdfWriter.addPage(pageObj)

    # save the Output in a file, wb for write binary
    pdfOutput = open(str(outputLocation +"/"+ outputFileName +".pdf"), 'wb')
    pdfWriter.write(pdfOutput)

    # close the pdfWriter post writing
    pdfOutput.close()


def moveMergePDF(sourcePath, destinationPath, filename):
    # move the output to the network destination folder
    dest_path = destinationPath
    source_path = sourcePath.replace('/','\\')
    file_name = "\\"+filename+".pdf"
    shutil.copyfile(source_path + file_name, dest_path + file_name)


if __name__ == "__main__":
    inputLocation  = str(input("Enter Input Location (Example: C:/MergePDF/Input) : "))
    outputLocation = str(input("Enter Output Location (Example: C:/MergePDF/Output) : "))
    outputFileName = str(input("Specify Output File Name : "))

    mergePDF(inputLocation, outputLocation, outputFileName)
    print("Output Merged File generated as :" + str(outputLocation + "/" + outputFileName + ".pdf"))

    moveFile = str(input("Do you want to move Merged file to an external Network Drive (Y/N) ? "))
    if (moveFile == 'Y'):
        # \\FTUATWVFMSAPP01\ADMP_DW_Production\Test Power Pivot\Saikat\
        driveLocation = str(input(r"Enter external Drive Location (Example: \\FTUATWVFMSAPP01\ADMP_DW_Production\Output) : "))
        moveMergePDF(outputLocation, driveLocation, outputFileName)
        print("Exiting Utility")
        exit(0)
    else:
        print("Exiting Utility")
        exit(0)