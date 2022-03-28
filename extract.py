
import win32com.client
from docx import Document
import os
import pandas as pd
import re



"""
    génère le docx pour un pdf donné
"""
def gen_docx(pdf, name_save):
    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0

    in_file = os.path.abspath(pdf)
    wb = word.Documents.Open(in_file)
    out_file = os.path.abspath("./gen_docx/" + name_save)
    print(out_file)
    wb.SaveAs2(out_file, FileFormat=16) # file format for docx
    wb.Close()



"""
    extrait les tables du docx
"""
def extract_tables(doc):
    doc = Document(doc)
    datalake = pd.DataFrame()
    dw = pd.DataFrame()
    tables = doc.tables
    for table in tables:
        data = [[cell.text for cell in row.cells] for row in table.rows]

        for x in range (len(data)):
            for y in range(len(data[0])):
                data[x][y] = re.sub(r'([0-9]) ([0-9])', r'\1\2', data[x][y])    # 1 000 -> 1000
                data[x][y] = re.sub(r'([0-9]),([0-9])', r'\1.\2', data[x][y])    # 1,00 -> 1.00
                data[x][y] = re.sub(r'([0-9]) %', r'\1%', data[x][y])    # 1,00 -> 1.00

        df = pd.DataFrame(data)

        # verifie si les premieres données de la dernière lignes ne sont pas suspects
        first_col = df.iloc[-1][0]
        sec_col = df.iloc[-1][1]

        if type(first_col) == str and type(sec_col) == str:
            if len(first_col) > 200 or len(sec_col) > 200 :
                df = df.iloc[:-1 , :]  # supprime la dernière ligne


        df = df.replace("", float("NaN"))
        df = df.dropna(how='all', axis=1)       # on retire les colonnes vides

        df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
        df = df.apply(pd.to_numeric, errors='ignore', downcast = 'float')
        datalake = pd.concat([datalake, df], ignore_index=True, sort=False)

        df = df.rename(columns={df.columns[0]: 'description'})
        dw = pd.concat([dw, df], ignore_index=True, sort=False)

        print(df, '\n\n')
    print(datalake)
    print(dw)
    # datalake.to_csv('./datalake.csv', index = False)
    datalake.to_excel('./datalake.xlsx')
    dw.to_excel('./dw.xlsx')

pdf = "./test.pdf"
# pdf = "./pdf_split/4 DECLARATION DE PERFORMANCE EXTRA-FINANCIERE.pdf"
gen_docx(pdf, "test.docx")
doc = './gen_docx/test.docx'
extract_tables(doc)