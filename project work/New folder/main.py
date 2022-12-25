from xml.etree.ElementTree import XML
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
from typing import Dict

import streamlit as st 

from PIL import Image

st.header("Form Filler Website")

#image = Image.open('1.png')
#st.image(image)



genre = st.radio(
    "What's your form format",
        ('None','Appointment letter', 'Pay letter', 'Extra letter'))

if genre == 'None':
    st.write('You selected None now.')

elif genre == 'Appointment letter':
    st.write('You selected appointment')

    uploadedFile = st.file_uploader('Upload kar jaldi', type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader")



    if st.button('Converter'):

        st.write('Shukriya dabane ke liye') #displayed when the button is clicked

        #my_bar = st.progress(0)

        #for percent_complete in range(100):
            #time.sleep(0.1)
            #my_bar.progress(percent_complete + 1)

       # with st.spinner('Wait for it...'):
            #time.sleep(5)
        st.success('Done!')    

        fname2=(r"D:\project work\PROJECCT.docx")

        os.chdir(r"D:\project work\saved_wordsfile")

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

        @st.cache(allow_output_mutation=True)
        def get_static_store() -> Dict:
            """This dictionary is initialized once and can be used to store the files uploaded"""
            return {}

        def file_selector(folder_path):
            filenames = os.listdir(folder_path)
            selected_filename = st.selectbox('Select a file', filenames)
            return os.path.join(folder_path, selected_filename)

        def main():
            fileslist = get_static_store()
            folderPath = 'D:\project work\saved_wordsfile'
            if folderPath:    
                filename = file_selector(folderPath)
                if not filename in fileslist.values():
                    fileslist[filename] = filename
            else:
                fileslist.clear()  # Hack to clear list if the user clears the cache and reloads the page
                st.info("Select one or more files.")

            #if st.button("Clear file list"):
                #fileslist.clear()
            #if st.checkbox("Show file list?", True):
                #finalNames = list(fileslist.keys())
                #st.write(list(fileslist.keys()))

            with open(filename, "rb") as fp:
                btn = st.download_button(
                    label="Download file",
                    data=fp,
                    file_name="myfile.docx",
                    mime="application/zip"
 )

        main()



      
        

    else:

        st.write('Convert ke liye upar dabaye') #displayed when the button is unclicked




elif genre == 'Pay letter':
    st.write('You selected Pay letter.')

elif genre == 'Extra letter':
    st.write('You selected Extra letter.')


else:
    st.write("You didn't select comedy.")







