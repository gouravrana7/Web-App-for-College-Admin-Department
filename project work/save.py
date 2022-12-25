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

from PIL import Image

st.header("Auto Form Filler Web App")




genre = st.sidebar.radio(
    "What's your form format",
        ('Choose Format Down Below','Appointment letter', 'Experience letter', 'Promotion Letter'))

if genre == 'Choose Format Down Below':
    
    image = Image.open('1.png')
    st.video("https://www.youtube.com/watch?v=BHKYD_prtZY&ab_channel=MM-DeemedtobeUniversity%27Official%27")
    st.image(image,width=500)

elif genre == 'Appointment letter':
    

    uploadedFile = st.file_uploader('Upload Excel/CSV file', type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader")



    if st.button('Converter'):

        st.write('Shukriya dabane ke liye') #displayed when the button is clicked

        my_bar = st.progress(0)

        for percent_complete in range(100):
            time.sleep(0.1)
            my_bar.progress(percent_complete + 1)

        with st.spinner('Wait for it...'):
            time.sleep(5)
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

        st.write("Download button down")  

    
        def zip_directory(folder_path, zip_path):
            with zipfile.ZipFile(zip_path, mode='w') as zipf:
                len_dir_path = len(folder_path)
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, file_path[len_dir_path:])
                    
        zip_directory('D:\project work\saved_wordsfile', 'D:\project work\saved_wordsfile.zip')

        a='D:\project work\saved_wordsfile.zip'

        with open(a, "rb") as fp:
            btn = st.download_button(
                label="Download ZIP",
                data=fp,
                file_name="myfile.zip",
                mime="application/zip"
            )



    else:

        st.write('Convert ke liye upar dabaye') #displayed when the button is unclicked




elif genre == 'Experience letter':
    
    option = st.radio(
    'Title',('Mr.','Ms.', 'Mrs.', 'Dr.'),horizontal=True)

    st.write('You selected:', option)

    a=st.date_input("Enter DATE")
    b=st.text_input("Enter NAME")
    c=st.text_input("Enter staff ID")
    d=st.text_input("Enter father name")
    e=st.text_input("Enter address")
    f=st.text_input("Enter city")
    g=st.text_input("Enter distt. / state")
    h = st.selectbox(
    'Select Designation',
    ('Professor and Head','Professor','Associate Professor','Assistant Professor'))
    st.write('You selected:', h)
    i=st.selectbox('Select Department',('Anatomy','Physiology','Bio-Chemistry','Pharmacology','Microbiology','Pathology','Forensic Medicine','Community Medicine','General Medicine','DVL','Psychiatry','Respiratory Medicine','Paediatrics','General Surgery','Orthopaedics','Ophthalmology','E.N.T.','Obst. & Gynae.','Anaesthesia','Radio-Diagnosis','Dentistry','Radiography','Physiotherapy','Para-Medical','Emergency Medicine',))
    j = st.selectbox(
    'Select College/Institute',
    ('MM College of Pharmacy','MM Engineering College','MM Institute of Management','MM Institute of Computer Technology and Business Management',
     'MM Institute of Computer Technology and Business Management (Hotel Management)','Department of Law','MM College of Dental Sciences and Research',
     'MM Institute of Medical Sciences and Research','MM College of Nursing','MM Institute of Nursing'))

    st.write('You selected:', j)
    k=st.date_input("Enter date from")
    l=st.date_input("Enter date to")
    m=st.date_input("Enter resination date")
    document_path = Path(__file__).parent / "office.docx"
    doc = DocxTemplate(document_path)

    
    if st.button('Submit'):
        
        context={"TITLE":option,"DATE":a,"NAME":b,"STAFFID":c,"FATHERNAME":d,"ADDRESS":e,"CITY":f,"DISTTSTATE":g,"DESIGNATION":h,"DEPARTMENT":i,"COLLEGE":j,"FROM":k,"TO":l,"RESINATIONDATE":m}

                    #new_list=[round(item,0)for item]

        doc.render(context)
        #doc.save('{}.docx'.format(b))
        doc.save(Path(__file__).parent/"generated_doc.docx")

        shutil.move('D:\\project work/generated_doc.docx','D:\\project work\\nee/generated_doc.docx')

        def zip_directory(folder_path, zip_path):
            with zipfile.ZipFile(zip_path, mode='w') as zipf:
                len_dir_path = len(folder_path)
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, file_path[len_dir_path:])
                    
        zip_directory('D:\\project work\\nee', 'D:\\project work\\nee.zip')

        ab='D:\\project work\\nee.zip'

        with open(ab, "rb") as ff:
            btn = st.download_button(
                label="Download ZIP",
                data=ff,
                file_name="mynewfile.zip",
                mime="application/zip"
            )

    else:
        st.write('Goodbye')

elif genre == 'Promotion Letter':
    

    option = st.radio(
        "Title",
        ('Mr.', 'Mrs.', 'Dr.','Ms.'),horizontal=True)

    a=st.text_input("Enter NAME")
    b=st.text_input("Enter staff ID")



    c = st.selectbox(
        'Select Present Designation',
        ('Professor','Associate Professor','Assistant Professor'))
    st.write('You selected:', c)

    n = st.selectbox(
        'Select Department',
        ('Biotechnology', 'Electronics and Communication Engineering','Electrical Engineering','Computer Science & Engineering', 'Mechanical Engineering', 'Mathematics Humanities', 'Chemistry','Civil Engineering', 'Physics'))
    st.write('You selected:', n)    

    d = st.selectbox(
        'Select College/Institute',
        ('MM College of Pharmacy','MM Engineering College','MM Institute of Management','MM Institute of Computer Technology and Business Management',
        'MM Institute of Computer Technology and Business Management (Hotel Management)','Department of Law','MM College of Dental Sciences and Research',
        'MM Institute of Medical Sciences and Research','MM College of Nursing','MM Institute of Nursing'))
    st.write('You selected:', d)


    e = st.selectbox(
        'Promoted As',
        ('Professor and Head','Professor','Associate Professor','Assistant Professor'))
    st.write('You selected:', e)


    f=st.date_input("DATE")
    g=st.text_input("CTC salary")

    document_path = Path(__file__).parent / "D:\project work\promotion.docx"
    doc = DocxTemplate(document_path)


    if st.button('Submit'):

        context={"NAME":a,"STAFFID":b,"PRESENTDESIGNATION":c,"DEPARTMENT":n,"COLLEGE/INSTITUTE":d,"PROMOTEDAS":e,"DATE":f,"CTCSALARY":g}

                    #new_list=[round(item,0)for item]

        doc.render(context)
        #doc.save('{}.docx'.format(a))
        doc.save(Path(__file__).parent/"promotion_letter.docx")

        shutil.move('D:\\project work/promotion_letter.docx','D:\\project work\\promotionLetter/promotion_letter.docx')

        def zip_directory(folder_path, zip_path):
            with zipfile.ZipFile(zip_path, mode='w') as zipf:
                len_dir_path = len(folder_path)
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, file_path[len_dir_path:])
                    
        zip_directory('D:\\project work\\promotionLetter', 'D:\\project work\\promotionLetter.zip')

        ab='D:\\project work\\promotionLetter.zip'

        with open(ab, "rb") as ff:
            btn = st.download_button(
                label="Download ZIP",
                data=ff,
                file_name="mynewfile.zip",
                mime="application/zip"
            )

    else:
        st.write('Goodbye')

else:
    st.write("You didn't select any.")  







