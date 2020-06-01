import sys
import locale
import ghostscript

outputFilePath="C:/MergePDF/Output/Sample.pdf"
content= "C:/MergePDF/Output/ConsolidatedGroupPack.pdf"
# "-dCompatibilityLevel=1.4",
# "-dPDFSETTINGS=/ebook",
args = [
    "ps2pdf", # actual value doesn't matter
    "-dNOPAUSE", "-dBATCH", "-dSAFER",
    "-sDEVICE=pdfwrite",
    "-dCompatibilityLevel=1.4",
    "-dPDFSETTINGS=/screen",
    "-dEmbedAllFonts=true",
    "-dSubsetFonts=true",
    "-dColorImageDownsampleType=/Bicubic",
    "-dColorImageResolution=110",
    "-dGrayImageDownsampleType=/Bicubic",
    "-dGrayImageResolution=110",
    "-sOutputFile=" +outputFilePath,
    "-c", ".setpdfwrite",
    "-f",  content
    ]
    # "sOutputFile=out.pdf",
    #  $1

# arguments have to be bytes, encode them
encoding = locale.getpreferredencoding()
args = [a.encode(encoding) for a in args]

print(args)
ghostscript.Ghostscript(*args)

