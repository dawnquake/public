import sys
import os
from docx2pdf import convert
from PyPDF2 import PdfFileMerger
import copy



os.chdir(os.path.dirname(os.path.realpath(__file__)))
part_names = ['title_page', 'executive_summary', 'main_text', 'references']

##Filter out the word documents in the current working directory
##Filter out the parts
##Will filter out the part which is last

filelist = os.listdir(os.path.dirname(os.path.realpath(__file__)))
dummy = []
word2pdf = []

#print(filelist)


for i in range(len(filelist)):
    filename = filelist[i]
    if filename.endswith('docx'):
        dummy.append(filename)

for j in range(len(part_names)):
    for i in range(len(dummy)):
        if dummy[i].startswith(part_names[j]):
            print(dummy[i])
            temp = copy.deepcopy(dummy[i])
    word2pdf.append(temp)

print(word2pdf)
to_be_deleted = []

for i in range(len(word2pdf)):
    print(word2pdf[i])
    convert(word2pdf[i])
    pdf_name = word2pdf[i].replace("docx", "pdf")
    to_be_deleted.append(os.path.join(os.path.dirname(os.path.realpath(__file__)), pdf_name))

print(to_be_deleted)



merger = PdfFileMerger()

for pdf in to_be_deleted:
    merger.append(pdf)

merger.write("3rd_year_Geophysics_Analysis.pdf")
merger.close()


for i in range(len(to_be_deleted)):
    os.remove(to_be_deleted[i])










