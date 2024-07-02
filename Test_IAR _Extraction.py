import pandas as pd
from PIL import Image
import pytesseract
import numpy as np
from pdf2image import convert_from_path, convert_from_bytes
from PyPDF2 import PdfReader
from PyPDF2 import PdfFileReader
from Crypto.Cipher import AES
import os
from datetime import datetime
from datetime import date
import re
import pandas as pd
import spacy
from spacy import displacy



def pdf2img(i):
    images = convert_from_path(r"D:\\UTA\\RA\\Inputs\\"+i, poppler_path=r"D:\Poppler\poppler-23.11.0\Library\bin")
    for img in range(len(images)):
        images[img].save(r'D:\\UTA\\RA\\Outputs\\temp_images\\'+str(i.split('.pdf')[0])+'_IMG_'+ str(img) +'.jpg', 'JPEG')
def keyW_search(filename,p):
    img1 = np.array(Image.open(filename))
    text = pytesseract.image_to_string(img1)
    paras = text.split('\n\n')
    for i in paras:
        if ("material weakness" in  i.lower()) :
            p.append(i)
    return(p)
def IAR_extraction(df):
    img_ = iter(os.listdir(r'D:\\UTA\\RA\\Outputs\\temp_images'))
    reprt = -1
    for i in range(len(os.listdir(r'D:\\UTA\\RA\\Outputs\\temp_images'))):
        pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
        pic = next(img_,-1)
        if pic!=-1:
            row_num = pic.split('_')[0]
            row_num = re.findall(r'\d+',row_num)[0]
            filename = r"D:\\UTA\\RA\\Outputs\\temp_images\\"+pic
            img = np.array(Image.open(filename))
            text = pytesseract.image_to_string(img)
            text = re.sub("independent auditor's report.*","independent auditors' report",text.lower())
            text = re.sub("independent auditor’s report.*","independent auditors' report",text.lower())
            text = re.sub("independent auditors’ report.*","independent auditors' report",text.lower())
            if (re.search(".*independent auditors' report.*",text.lower())) :
                reprt = 1
                fh = open(r'D:\\UTA\\RA\\Outputs\\'+str('_'.join(pic.split('_')[:-2]))+'.txt','a')
                fh.writelines(text)
                fh.write('\n\n\n')
                for i in range(0,4,1):
                    iar_img = next(img_,-1)
                    if iar_img != -1:
                        img_1 = np.array(Image.open(r"D:\\UTA\\RA\\Outputs\\temp_images\\"+iar_img))
                        t_1 = pytesseract.image_to_string(img_1)
                        reprt = 1
                        fh = open(r'D:\\UTA\\RA\\Outputs\\'+str('_'.join(pic.split('_')[:-2]))+'.txt','a')
                        fh.writelines(t_1)
                        fh.write('\n\n\n')
                        fh.close()
            elif (re.search("independent auditor's report.*",text.lower())) :
                reprt = 1
                fh = open(r'D:\\UTA\\RA\\Outputs\\'+str('_'.join(pic.split('_')[:-2]))+'.txt','a')
                fh.writelines(text)
                fh.write('\n\n\n')
                for i in range(0,4,1):
                    iar_img = next(img_,-1)
                    if iar_img != -1:
                        img_1 = np.array(Image.open(r"D:\\UTA\\RA\\Outputs\\temp_images\\"+iar_img))
                        t_1 = pytesseract.image_to_string(img_1)
                        reprt = 1
                        fh = open(r'D:\\UTA\\RA\\Outputs\\'+str('_'.join(pic.split('_')[:-2]))+'.txt','a')
                        fh.writelines(t_1)
                        fh.write('\n\n\n')
                        fh.close()
        else:
            continue
    if reprt == -1:
        df.loc[int(row_num),'auditor_report'] = '0'
    elif reprt == 1:
        df.loc[int(row_num),'auditor_report'] = '1'
def spaCy_ner(pdf):
    d = []
    m = []
    l = []
    date =''
    money =''
    loc =''
    NER = spacy.load("en_core_web_sm")
    reader_i = PdfReader(r"D:\\UTA\\RA\\Inputs\\"+pdf)
    page_1_i = reader_i.pages[0].extract_text()
    pg_1_i = page_1_i.replace('\n','')
    ner = NER(pg_1_i)
    for i in ner.ents:
        if i.label_ == 'DATE':
            d.append(i.text)
        elif i.label_ == 'MONEY':
            m.append(i.text)
        elif i.label_ == 'GPE':
            l.append(i.text)
    if d:
        date = d[0]
    if m:
        money = m[0]
    if l:
        loc = l[0]
    return date,money,loc
def sort_on_first4(item):
    res = item.split('_')[1]
    return res
def search_extract(i):
    df = pd.read_excel(r'D:\UTA\RA\Outputs\MW Entities Unique Cusip_09122023_v2_TEST.xlsx')
    for files in list(os.listdir(r"D:\\UTA\\RA\\Outputs\\temp_images\\")):
        os.remove(r"D:\\UTA\\RA\\Outputs\\temp_images\\"+files)
    pdf2img(i)
    p=[]
    for i in os.listdir(r'D:\\UTA\\RA\\Outputs\\temp_images'):
        if i.endswith(".jpg"):
            pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
            filename = r"D:\\UTA\\RA\\Outputs\\temp_images\\"+i
            paras = keyW_search(filename,p)
            row_num = i.split('_')[0]
            row_num = re.findall(r'\d+',row_num)[0]
            if len(paras) == 0: 
                df.loc[int(row_num),'disclose or not'] = '0'
            else:
                df.loc[int(row_num),'disclose or not'] = '1'
                df.loc[int(row_num),'TEXT'] = '\n'.join(paras)
    IAR_extraction(df)
    for files in list(os.listdir(r"D:\\UTA\\RA\\Outputs\\temp_images\\")):
        os.remove(r"D:\\UTA\\RA\\Outputs\\temp_images\\"+files)
    df.to_excel(r"D:\UTA\RA\Outputs\MW Entities Unique Cusip_09122023_v2_TEST.xlsx",index=False)

def flag_dupes(dupe_l):
    df_ = pd.read_excel(r'D:\UTA\RA\Outputs\MW Entities Unique Cusip_09122023_v2_TEST.xlsx')
    df = df_.copy()
    proc_pdf = dupe_l[-1]
    row_num = proc_pdf.split('_')[0]
    row_num = re.findall(r'\d+',row_num)[0]
    mw_flag = df.loc[int(row_num),'disclose or not']
    mw_txt = df.loc[int(row_num),'TEXT']
    iar_flag_ = df.loc[int(row_num),'auditor_report']
    if iar_flag_ == '0':
        iar_flag = iar_flag_
    else:
        iar_flag = 'Same as: '+proc_pdfT
    for pdf in dupe_l[:-1]:
        d_row_num = pdf.split('_')[0]
        d_row_num = re.findall(r'\d+',d_row_num)[0]
        df.loc[int(d_row_num),'disclose or not'] = mw_flag
        df.loc[int(d_row_num),'TEXT'] = mw_txt 
        df.loc[int(d_row_num),'auditor_report'] = iar_flag
    df.to_excel(r"D:\UTA\RA\Outputs\MW Entities Unique Cusip_09122023_v2_TEST.xlsx",index=False)


now_ = datetime.now().strftime("%H:%M:%S")
date_ = date.today().strftime("%b-%d-%Y")
log_file_path=r"D:\UTA\RA\Outputs\\File_Generation_Log.txt"
log_file=open(log_file_path,mode='a',newline='\n')
print(f"\n\n\nStart Date:{date_}\t\tStart time: {now_}",file=log_file)
list_file = []
enc_file = []
unique_list_file = []
dupe_l = []
for i in os.listdir(r'D:\UTA\RA\Inputs'):
    list_file.append(i)
    list_file.sort(key=sort_on_first4) 
for i in range(len(os.listdir(r'D:\UTA\RA\Inputs'))):
    now = datetime.now().strftime("%H:%M:%S")
    try:
        j = i+1
        if j < len(os.listdir(r'D:\UTA\RA\Inputs')):
            d_i, m_i, l_i = spaCy_ner(list_file[i])
            d_j, m_j, l_j = spaCy_ner(list_file[j])
            if ((re.findall(r'(?<=_)[^_]+',list_file[i])[1] == re.findall(r'(?<=_)[^_]+',list_file[j])[1]) and (d_i == d_j) and (m_i == m_j) and (l_i==l_j)):
                print(f"\n{list_file[i]} == {list_file[j]}\n{d_i, m_i, l_i}=={d_j, m_j, l_j}",file=log_file)
                dupe_l.append(list_file[i])
                dupe_l.append(list_file[j])
                dupe_l = sorted(list(set(dupe_l)))
            else:
                print(f"\n{dupe_l}",file=log_file)
                if len(dupe_l) > 0:
                    flag_dupes(dupe_l)
                    dupe_l = []
                print(f"\n{list_file[i]} !! {list_file[j]}\n{d_i, m_i, l_i} !! {d_j, m_j, l_j}",file=log_file)
                unique_list_file.append(list_file[i])
                search_extract(list_file[i])
        else:
            print(f"\nLast File: {list_file[j-1]}",file=log_file)
            unique_list_file.append(list_file[j-1])
            search_extract(list_file[i])
        end = datetime.now().strftime("%H:%M:%S")
        print(f"Time Take for completion->{datetime.strptime(end, '%H:%M:%S') - datetime.strptime(now, '%H:%M:%S')}",file=log_file)
    except Exception as e:
        print(f"\nError:-> {list_file[i]} has an error {e}",file=log_file)
        if len(dupe_l) > 0:
            search_extract(dupe_l[-1])
            flag_dupes(dupe_l)
            dupe_l = []
end_ = datetime.now().strftime("%H:%M:%S")
end_date_ = date.today().strftime("%b-%d-%Y")
print(f"\n\n\nEnd Date:{end_date_}\t\tEnd time: {end_}\nTime Take for completion->{datetime.strptime(end_, '%H:%M:%S') - datetime.strptime(now_, '%H:%M:%S')}",file=log_file)
log_file.close()
