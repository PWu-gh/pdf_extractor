import fitz
import unidecode 
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTChar, LTTextBox
from collections import Counter
import re


"""
    retourne la page et le label où se trouve le sommaire
"""
def loc_sommaire(doc):
    page_id = 0
    for page in doc:
        page_id += 1
        txt = page.get_text()[:200].upper()
        if "SOMMAIRE" in txt:
            # print(page.get_label())
            return page_id, page.get_label()
        


"""
    extrait les textes et la taille des textes avec pdfminer
"""
def extract_fontsize(path, page_id):
    Extract_Data=[]
    doc = extract_pages(path)
    for page_layout in doc:
        if page_layout.pageid == page_id:
            for element in page_layout:
                if isinstance(element, LTTextBox):
                    for text_line in element:
                        for character in text_line:
                            if isinstance(character, LTChar):
                                Font_size = int(character.size)
                    text = unidecode.unidecode(element.get_text())
                    text = text.replace('\n', '')
                    Extract_Data.append([Font_size, text])
            break
    
    # reconnect strings
    corrected_data = []
    if len(Extract_Data) > 0:
        last_string_size = 0
        for data in Extract_Data:
            if last_string_size == data[0]:
                if corrected_data[-1][1][-1:] == ' ':
                    corrected_data[-1][1] = corrected_data[-1][1] + data[1]
            else:
                corrected_data.append(data)
            last_string_size = data[0]

    return corrected_data


"""
    retourne les textes mis en valeur (taille plus grand que la police normale), 
    on considère que c'est le titre
"""
def get_titles(data):
    tab_size = [col[0] for col in data]
    collec_size = dict(Counter(tab_size))
    most_common_size = max(collec_size, key=collec_size.get)

    for i in collec_size.copy():
        if i < most_common_size:
            del collec_size[i]
        elif collec_size[i] < 5:
            del collec_size[i]

    if len(collec_size) >= 2:
        del collec_size[most_common_size]
    title_size = max(collec_size, key=collec_size.get)
        
    titles = []
    for x in data:
        if x[0] == title_size:
            if x[1].isdigit() == False:
                # titles.append(x)
                titles.append(x[1].strip())
    return titles


"""
    utilise fitz de pymuPDF pour mettre les textes dans un tableau
"""
def pdfto_array_fitz(pdf):
    document  = fitz.open(pdf) 
    fitz_txt = []
    for page in document:
        str = unidecode.unidecode(page.get_text())
        fitz_txt = str.split('\n')
    return fitz_txt


"""
    génère et retourne les parties du sommaire dans un tableau
"""
def gen_sommaire(document, som_page):
    # récupère le texte du sommaire
    txt = document[som_page].get_text()
    txt = unidecode.unidecode(txt)
    txt = txt.replace('\n', ' ').replace('  ', ' ')

    # titre , page
    sommaire = []
    for title in titles:
        # g1 = titre
        # g2 = annee true -> page
        # g3 = annee false -> page
        reg_page = r'(' + title + r')' + r' ([1-9][0-9]*) ?([0-9]*)'
        tx = re.search(reg_page, txt)
        # print(tx.group(3), tx.group(0))
        if tx.group(3) == '':
            sommaire.append([tx.group(1), int(tx.group(2))])
        else:
            sommaire.append([tx.group(1)+" "+tx.group(2), int(tx.group(3))])
    return sommaire


"""
    génère les pdf pour chaque parties du sommaire
"""
def gen_pdf(sommaire, document):
    # file_name = './pdf_split/test.pdf'
    for x in range(len(sommaire)):
        from_p = sommaire[x][1]+1       # 2 premieres pages ne comptent pas
        if x == (len(sommaire)-1):
            to_p = document.page_count
        else:
            to_p = sommaire[x+1][1]

        doc2 = fitz.open()
        doc2.insert_pdf(document, from_page = from_p, to_page = to_p, start_at = 0)
        doc2.save("./pdf_split/"+ sommaire[x][0] +".pdf")




pdf = "./U_record/TF1_2020.pdf"
document  = fitz.open(pdf) 

som_page = loc_sommaire(document)[0]
tab_fontsize = extract_fontsize(pdf, som_page)
titles = get_titles(tab_fontsize)
sommaire = gen_sommaire(document, som_page-1)
gen_pdf(sommaire, document)


# doc2 = fitz.open()
# doc2.insert_pdf(document, from_page = 141, to_page = 143, start_at = 0)
# doc2.save("./pdf_split/"+ "test" +".pdf")