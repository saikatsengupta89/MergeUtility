import os, shutil, glob, img2pdf
from pdf2jpg import pdf2jpg
from PIL import Image

# def compress( pdf_files_folder, new_pdf_file):
#
#     files = os.listdir(r"{}\{}".format(current_location='C:/MergePDF/Output/', pdf_files_folder='C:/MergePDF/Output/'))
#     files = [f for f in files if f.endswith(".pdf")]
#     folders = r"{}\images".format(pdf_files_folder)
#     images = []
#
#     for file in files:
#         pdf2jpg.convert_pdf2jpg(r"{}\{}".format( pdf_files_folder, file), dpi=600, pages="ALL")
#
#         for folder in os.listdir(folders):
#             location = r"{}\{}".format(folders, folder)
#
#             for image in os.listdir(location):
#                 full_location = r"{}\{}".format(location, image)
#                 images.append(full_location)
#
#         with open(new_pdf_file,"wb") as f:
#             f.write(img2pdf.convert(images))
#
#     shutil.rmtree(folders)
#
# if __name__=="__main__":
#     compress('C:/MergePDF/Output', 'C:/MergePDF/Output/Output.pdf')



files = os.listdir(r"C:/MergePDF/Output/")
files = [f for f in files if f.endswith(".pdf")]
print(files)
folders = r"{}/images".format("C:/MergePDF/Output")
images=[]
new_pdf_file= "C:/MergePDF/Output/ConsolidatedGroupPack_comp.pdf"
print(folders)

pdf2jpg.convert_pdf2jpg(r"C:/MergePDF/Output/ConsolidatedGroupPack.pdf", r"C:/MergePDF/Output/Images",  dpi=80, pages="ALL")
for image in os.listdir("C:/MergePDF/Output/Images/ConsolidatedGroupPack.pdf_dir"):
    full_location= r"{}/{}".format("C:/MergePDF/Output/Images/ConsolidatedGroupPack.pdf_dir", image)
    print(full_location)
    images.append(full_location)

with open(new_pdf_file, "wb") as f:
    f.write(img2pdf.convert(images))

f.close()