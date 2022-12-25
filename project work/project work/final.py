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

st.header("Auto Form Filler Web App")




genre = st.sidebar.radio(
    "What's your form format",
        ('Choose Format Down Below','Appointment letter', 'Experience letter', 'Promotion Letter',
        'Contact Engineer For Help','Feedback'))

if genre == 'Choose Format Down Below':
    
    
    st.video("https://www.youtube.com/watch?v=BHKYD_prtZY&ab_channel=MM-DeemedtobeUniversity%27Official%27")
   
    col1, col2 = st.columns(2)
    image = Image.open('rana.jpg')
    image = Image.open('saini.jpeg')
    with col1:
        st.header("Gourav Rana \n Btech Cse 7th")
        st.image("rana.jpg")

    with col2:
        st.header("Neeraj Saini \n Btech Cse 7th")
        st.image("saini.jpeg",width=280)

elif genre == 'Appointment letter':
    

    uploadedFile = st.file_uploader('Upload Excel/CSV file', type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader")



    if st.button('Converter'):

        st.write('Shukriya dabane ke liye') #displayed when the button is clicked

        my_bar = st.progress(0)

        for percent_complete in range(100):
            time.sleep(0.1)
            my_bar.progress(percent_complete + 1)

        with st.spinner('Wait for it...'):
            time.sleep(2)
        st.success('Done!')    

        fname2=(r"D:\college\project work\PROJECCT.docx")

        os.chdir(r"D:\college\project work\saved_wordsfile")
        



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
                    
        zip_directory('D:\college\project work\saved_wordsfile', 'D:\college\project work\saved_wordsfile.zip')

        a='D:\college\project work\saved_wordsfile.zip'

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

        shutil.move('D:\college\project work/generated_doc.docx','D:\college\project work\\nee/generated_doc.docx')

        def zip_directory(folder_path, zip_path):
            with zipfile.ZipFile(zip_path, mode='w') as zipf:
                len_dir_path = len(folder_path)
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, file_path[len_dir_path:])
                    
        zip_directory('D:\college\project work\\nee', 'D:\college\project work\\nee.zip')

        ab='D:\college\project work\\nee.zip'

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

    document_path = Path(__file__).parent / "D:\college\project work\promotion.docx"
    doc = DocxTemplate(document_path)


    if st.button('Submit'):

        context={"NAME":a,"STAFFID":b,"PRESENTDESIGNATION":c,"DEPARTMENT":n,"COLLEGE/INSTITUTE":d,"PROMOTEDAS":e,"DATE":f,"CTCSALARY":g}

                    #new_list=[round(item,0)for item]

        doc.render(context)
        #doc.save('{}.docx'.format(a))
        doc.save(Path(__file__).parent/"promotion_letter.docx")

        shutil.move('D:\college\project work\\promotion_letter.docx','D:\college\project work\\promotionLetter/promotion_letter.docx')

        def zip_directory(folder_path, zip_path):
            with zipfile.ZipFile(zip_path, mode='w') as zipf:
                len_dir_path = len(folder_path)
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, file_path[len_dir_path:])
                    
        zip_directory('D:\college\project work\\promotionLetter', 'D:\college\project work\\promotionLetter.zip')

        ab='D:\college\project work\\promotionLetter.zip'

        with open(ab, "rb") as ff:
            btn = st.download_button(
                label="Download ZIP",
                data=ff,
                file_name="mynewfile.zip",
                mime="application/zip"
            )

    else:
        st.write('Goodbye')




elif genre == 'Contact Engineer For Help':
    st.write("Click down for to contact 'Gourav Rana or Neeraj Saini'")

    if st.button("Email ID of Er 'Gourav Rana'"):
        st.write('gourav.rana70000@gmail.com')
    else:
        st.write('')

    if st.button("Email ID of Er 'Neeraj Saini'"):
        st.write('neerajsaini2k1@gmail.com')
    else:
        st.write('')




elif genre == 'Feedback':
    import streamlit as st
    def create_table():
        c.execute('CREATE TABLE IF NOT EXISTS feedback(date_submitted DATE, Q1 TEXT,  Q3 INTEGER, Q6 TEXT,  Q8 TEXT)')

    def add_feedback(date_submitted, Q1, Q3, Q6, Q8):
        c.execute('INSERT INTO feedback (date_submitted,Q1, Q3, Q6, Q8) VALUES (?,?,?,?,?)',(date_submitted,Q1, Q3,  Q6,  Q8))
        conn.commit()

    def main():

        st.title("Feedback")

        d = st.date_input("Today's date",None, None, None, None)
        
        question_1 = st.selectbox('How you feel our Web App?',('','Excellent', 'Very Good', 'Good','Average','Poor','Very poor'))
        #st.write('You selected:', question_1)
        
        #question_2 = st.slider('What age are you in?', 0,50)
        #st.write('You selected:', question_2) 
        
        question_3 = st.slider('Overall, how happy are you with the Form format? (5 being very happy and 1 being very dissapointed)', 1,5,1)
        #st.write('You selected:', question_3)

        #question_4 = st.selectbox('Web App is fast or not',('','Yes', 'No'))
        #st.write('You selected:', question_4)

        #question_5 = st.selectbox('Does all options is working correctly?',('','Yes', 'No'))
        #st.write('You selected:', question_5)
        question_6 = st.radio(
        'Does it need improvement',('Yes', 'No'),horizontal=True)
        #question_6 = st.selectbox('Does it need improvement',('','Yes', 'No'))
        #st.write('You selected:', question_6)

        #question_7 = st.selectbox('We should use different themes',('','Yes', 'No'))
        #st.write('You selected:', question_7)

        question_8 = st.text_input('What could have been better?', max_chars=50)

        if st.button("Submit feedback"):
            create_table()
            add_feedback(d, question_1,  question_3, question_6, question_8)
            st.success("Feedback submitted")

    if __name__ == '__main__':
        main()        

else:
    st.write("You didn't select any.")








