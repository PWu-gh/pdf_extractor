from part_pdf import *
from extract import *


pdf = "./U_record/TF1_2020.pdf"
document  = fitz.open(pdf) 

som_page = loc_sommaire(document)[0]
tab_fontsize = extract_fontsize(pdf, som_page)
titles = get_titles(tab_fontsize)
sommaire = gen_sommaire(document, som_page-1, titles)
gen_pdf(sommaire, document)


pdf = "./test.pdf"
# pdf = "./pdf_split/4 DECLARATION DE PERFORMANCE EXTRA-FINANCIERE.pdf"
gen_docx(pdf, "test.docx")
doc = './gen_docx/test.docx'
extract_tables(doc)