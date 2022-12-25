from turtle import width
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
import numpy as np
import pandas as pd
import base64
import sqlite3
conn = sqlite3.connect('student_feedback.db')
c = conn.cursor() 

from PIL import Image

uploadedFile = st.file_uploader('Upload Excel/CSV file', type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader")



if st.button('Converter'):

    st.write('Shukriya dabane ke liye') #displayed when the button is clicked

    #my_bar = st.progress(0)

    '''for percent_complete in range(100):
        time.sleep(0.1)
        my_bar.progress(percent_complete + 1)

    with st.spinner('Wait for it...'):
        time.sleep(5)
    st.success('Done!')  '''  

    fname2=(r"E:\\project work\\PROJECCT.docx")

    os.chdir(r"E:\\project work\\saved_wordsfile")

    df=pd.read_excel(uploadedFile)

    #Date=df["Date"].values
    NAME=df["NAME"].values
    FNAME=df["FNAME"].values
    ROW3=df["ROW3"].values
    ROW4=df["ROW4"].values
    DESIGNATION=df["DESIGNATION"].values
    COLLEGEINSTITUTE=df["COLLEGEINSTITUTE"].values
    BASICPAY=df["BASICPAY"].values
    TA=df["TA"].values
    MA=df["MA"].values
    PP=df["PP"].values
    ATOTAL=df["ATOTAL"].values
    GRATUITY=df["GRATUITY"].values
    ABTOTAL=df["ABTOTAL"].values
    EPF=df["EPF"].values
    ESI=df["ESI"].values
    EPFESIG=df["EPFESIG"].values
    #Joining_Date=df["Joining_Date"].values
    #Posting=df["Posting"].values
    #Supervisor_Name=df["Supervisor_Name"].values


    zipped=zip(NAME,FNAME,ROW3,ROW4,DESIGNATION,COLLEGEINSTITUTE,BASICPAY,TA,MA,PP,ATOTAL,GRATUITY,ABTOTAL,EPF,ESI,EPFESIG)

    for a,b,c,d,e,f,j,n,o,r,s,t,u,w,y,z in zipped:

        doc=DocxTemplate(fname2)

        context={"NAME":a,"FNAME":b,"ROW3":c,"ROW4":d,"DESIGNATION":e,"COLLEGEINSTITUTE":f,"BASICPAY":j,"TA":n,"MA":o,"PP":r,"ATOTAL":s,"GRATUITY":t,"ABTOTAL":u,"EPF":w,"ESI":y,"EPFESIG":z}

        #new_list=[round(item,0)for item]

        doc.render(context)
        doc.save('{}.docx'.format(a))

    st.write("Download button down")  


    def zip_directory(folder_path, zip_path):
        with zipfile.ZipFile(zip_path, mode='w') as zipf:
            len_dir_path = len(folder_path)
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, file_path[len_dir_path:])
                
    zip_directory('E:\\project work\\saved_wordsfile', 'E:\\project work\\saved_wordsfile.zip')

    a='E:\\project work\\saved_wordsfile.zip'

    with open(a, "rb") as fp:
        btn = st.download_button(
            label="Download ZIP",
            data=fp,
            file_name="myfile.zip",
            mime="application/zip"
        )

else:

    st.write('Convert ke liye upar dabaye') #displayed when the button is unclicked