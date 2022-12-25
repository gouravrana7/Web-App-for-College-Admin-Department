import pandas as pd
import os
from docxtpl import DocxTemplate
import streamlit as st 
import pandas as pd
import os
import sys
from docxtpl import DocxTemplate
import time
import base64
import shutil
import zipfile
from os.path import basename
from zipfile import ZipFile
from io import BytesIO
from pathlib import Path
import streamlit as st 
from docxtpl import DocxTemplate
import streamlit_authenticator as stauth
import os.path
import shutil
import streamlit as st 
import streamlit.components.v1 as components

uploadedFile = st.file_uploader('Upload Excel/CSV file', type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader")

if st.button('Converter'):

        st.write('Shukriya dabane ke liye') #displayed when the button is clicked

        my_bar = st.progress(0)

        #for percent_complete in range(100):
            #time.sleep(0.1)
            #my_bar.progress(percent_complete + 1)

        #with st.spinner('Wait for it...'):
            #time.sleep(5)
        st.success('Done!')
        fname2=(r"D:\\project work\\expmailmerging.docx")

        os.chdir(r"D:\\project work\\neeju")

        df=pd.read_excel(uploadedFile)


        NAME=df["NAME"].values
        FATHERNAME=df["FATHERNAME"].values
        ADDRESS=df["ADDRESS"].values
        DISTTSTATE=df["DISTTSTATE"].values
        STAFFID=df["STAFFID"].values
        DESIGNATION=df["DESIGNATION"].values
        DEPARTMENT=df["DEPARTMENT"].values
        COLLEGE=df["COLLEGE"].values
        FROM=df["FROM"].values
        TO=df["TO"].values
        RESIGNATIONDATE=df["RESIGNATIONDATE"].values
        DATE=df["DATE"].values


        zipped=zip(NAME,FATHERNAME,ADDRESS,DISTTSTATE,STAFFID,DESIGNATION,DEPARTMENT,COLLEGE,FROM,TO,RESIGNATIONDATE,DATE)

        for b,c,d,e,f,g,h,i,j,k,l,m in zipped:

            doc=DocxTemplate(fname2)
            context={"NAME":b,"FATHERNAME":c,"ADDRESS":d,"DISTTSTATE":e,"STAFFID":f,"DESIGNATION":g,"DEPARTMENT":h,"COLLEGE":i,"FROM":j,"TO":k,"RESIGNATIONDATE":l,"DATE":m}

            #new_list=[round(item,0)for item]

            doc.render(context)
            adda=doc.save('{}.docx'.format(b))
            #doc.save(Path(__file__).parent/"newFile.docx")
            completeName = os.path.join("D:\project work\neeju", "adda") 

            HtmlFile = open("D:\project work\Default.htm", 'r', encoding='utf-8')
            source_code = HtmlFile.read() 
            print(source_code)
            components.html(source_code)


   