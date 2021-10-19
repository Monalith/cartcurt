import base64
from sys import version
import xlrd
import csv, numpy
print(version)
file_Name = ("vcarddata/liste.xlsx")
wb = xlrd.open_workbook(file_Name)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
nof_Rows=sheet.nrows
data_Excel=sheet.row_values(1)
i=1
while i<nof_Rows:
    data_In=data_Excel=sheet.row_values(i)      
    card_Name=str(data_In[0])
    card_Surname=str(data_In[1])
    card_mTitle=str(data_In[2])
    card_Bday=str(data_In[3])
    card_Org=str(data_In[5])
    card_Title=str(data_In[6])
    card_Tel=str(data_In[4])
    card_Email=str(data_In[7])
    card_Url=str(data_In[8])
    card_Adress=str(data_In[9])
    card_Filen=str(data_In[10])
    print(data_In)
    with open(str('vcarddata/'+str(card_Filen[:15])+'.jpg'), "rb") as img_file:
        my_string = base64.b64encode(img_file.read())
    #print('buradan başlıyor:',my_string.decode('utf-8'))
    image=my_string.decode('utf-8')
    image_Data=['PHOTO;ENCODING=BASE64;TYPE=PNG:'+image]
    line1 = ['BEGIN:VCARD','VERSION:3.0','REV:2021-05-28T00:00:00Z','N:'+card_Surname+';'+card_Name+' ;;'+card_mTitle+';','BDAY:'+card_Bday,'ORG:'+card_Org,'TITLE:'+card_Title,'TEL;CELL:'+card_Tel,'EMAIL;INTERNET;WORK:'+card_Email,'URL;WORK:'+card_Url,'ADR;WORK:'+card_Adress]
    line2=['END:VCARD']
    def save_vcf(data):
        with open(str('vcards/'+str(card_Filen[:15])+'.vcf'),'a') as file:
            csvWriter = csv.writer(file, delimiter=',')       
            #print(data)
            r=0
            for rows in data:
                #print(str(data[r]))
                csvWriter.writerow([str(data[r])])
                r=r+1
            
    save_vcf(line1)
    save_vcf(image_Data)
    save_vcf(line2)
    i=i+1
