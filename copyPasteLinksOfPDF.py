from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import AnnotationBuilder,Fit
import os
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import sys

def linkCreator(PyPDF2WriterObject:PdfWriter, linkLocationPage:int, linkBoxSize:tuple, gotoPage:int):
    linkLocationPage = linkLocationPage - 1  # zero-indexed page
    gotoPage = gotoPage - 1  # zero-indexed page
    annotation = AnnotationBuilder.link(rect=linkBoxSize, target_page_index=gotoPage, fit=Fit("/Fit"))
    PyPDF2WriterObject.add_annotation(page_number=linkLocationPage, annotation=annotation)

def copyPasteLinksofPDF(sourcePDF:str,destinationPDF:str)->bool:
    if not (os.path.exists(sourcePDF) and os.path.exists(destinationPDF)):
        print(f"The PDF is not found"
                  f"source PDF exists ? : {os.path.exists(sourcePDF)}"
                  f"Destination PDF exists ?: {os.path.exists(destinationPDF)}")
        return False
    pageNumber = 1
    reader = PdfReader(sourcePDF)
    reader2 = PdfReader(destinationPDF)
    writer = PdfWriter()
    rectangleDimensions= []
    for page in reader2.pages:#adding all pages of destination PDF to destination pdf why? then only it works
        writer.add_page(page) #this blocks puts all pages in output pdf
    
    for page in reader.pages:#getting all objects from source pdf
        if "/Annots" in page:
            for annot in page["/Annots"]: # type: ignore
                obj = annot.get_object()
                subtype = annot.get_object()["/Subtype"]
                if subtype == "/Link":
                    rectangleDimensions.append(obj["/Rect"])
                    #obj2 = obj.get_object()["/A"] #3 lines for digging and getting id of an dictionary variable lies beneath page object
                    #obj3 = obj2.get_object()["/D"]#the destination page changes when target page of link changes with diff of 28
                    #obj4 = obj3.get_object()[0]#page index starts from 0 so -28-1=29 -1 done in above command
                    # this method not working if pages go above 9
                    # Accessing the destination page number
                    #destination_page = obj4.idnum

            rectangleDimensions.sort()
            for destination_page,dimension in enumerate(rectangleDimensions):
                linkCreator(writer,linkLocationPage=pageNumber,linkBoxSize=dimension,
                            gotoPage=destination_page+1)
            rectangleDimensions.clear()    
            pageNumber+=1
    with open(destinationPDF, "wb") as fp:
        writer.write(fp)
    print("PDF links copied successfully.")
    return True

if __name__=="__main__": # for passing argument in powershell
    """import argparse
    parser=argparse.ArgumentParser()
    parser.add_argument("--Template_PDF_location",type=str,required=True)
    parser.add_argument("--Output_PDF_location",type=str,required=True)
    args = parser.parse_args()

    copyPasteInternalLinksfromPDF(args.Template_PDF_location,args.Output_PDF_location)  """

    copyPasteLinksofPDF(r"C:\Users\Bala krishnan\OneDrive\Documents\Python projects\copy links from one pdf to another\source.pdf",
                              r"C:\Users\Bala krishnan\OneDrive\Documents\Python projects\copy links from one pdf to another\destination.pdf") 