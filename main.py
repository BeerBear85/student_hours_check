
import argparse
import pandas as pd
import os
from PyPDF2 import PdfReader
import re

#read argument for input file folder
parser = argparse.ArgumentParser()
parser.add_argument("--input_folder", default="input_pdf_files", help="input file folder")
parser.add_argument("--department_name", default="EEALD", help="ie. EEALD")
parser.add_argument("--person_excel_file", default="eeal_people.xlsx", help="File from confluence with person data")
args = parser.parse_args()

#create abs path to person excel file
args.person_excel_file = os.path.join(os.getcwd(), args.person_excel_file)


#import excel file (the 'ECS_Allocation' sheet) with person data into pandas data frame
# check that the file exists and is available for reading
try:
    open(args.person_excel_file, 'r')
except IOError:
    print("File not accessible")
else:
    df_persons = pd.read_excel(args.person_excel_file, sheet_name='ECS_Allocation')

#Drop entries not matching the department name
df_persons = df_persons[df_persons['Department'] == args.department_name]

#For each pdf file in the input folder, extract text and store in array
pdf_text_array = []
for filename in os.listdir(args.input_folder):
    if filename.endswith(".pdf"):
        #create abs path to pdf file
        filename = os.path.join(os.getcwd(), args.input_folder, filename)
        #extract text from pdf file using PdfReader
        with open(filename, 'rb') as f:
            pdf = PdfReader(f)
            #get number of pages in pdf file
            num_pages = len(pdf.pages)
            #extract text from all pages
            pdf_text = ''
            for page_number in range(num_pages):
                pdf_page_text = pdf.pages[page_number].extract_text()
                pdf_text += pdf_page_text
            #append text to array
            pdf_text_array.append(pdf_text)
    else:
        continue

#For each PDF file, create dict with key values
key_value_array = []
for file_text in pdf_text_array:
    
    key_value_dict = {}

    #use regex to find the person name (text after "Medarbejdernavn" in same line)
    person_name = re.search(r'Medarbejdernavn: (.*)', file_text)
    key_value_dict['person_name'] = person_name[1].strip() #get the match in group 1

    #use regex to find the person number (text after "Medarbejdernr." in same line)
    person_number = re.search(r'Medarbejdernr.: (.*)', file_text)
    key_value_dict['person_number'] = person_number[1].strip()

    try:
        #use regex to find all the total time (text after "Totaltid" in same line)
        total_time = re.findall(r'Totaltid (.*)', file_text)
        #extract the last match
        total_time = total_time[-1]
        #use regex to extract and convert the time in the end of total time (the XX:XX) in the end of the match
        total_time = re.search(r'(\d+:\d{2})', total_time)
        total_time = total_time[total_time.lastindex].strip()
    #catch exception for TypeError if no match is found (no hours registred)
    #catch exception for IndexError if no match is found (no hours registred)
    except (TypeError, IndexError):
        total_time = '0:00'
    key_value_dict['total_time'] = total_time

    key_value_array.append(key_value_dict)


# Create email reply text for each person matching the department
full_email_string = 'Følgende tidsregistreringer for ' + args.department_name + ' er godkendt:\n\n'
for timesheet_key_value in key_value_array:
    #Check if name of timesheet matches name in person excel file
    for index, row in df_persons.iterrows():
        #print('Matching ' + timesheet_key_value['person_name'] + ' with ' + row['Name'])
        if timesheet_key_value['person_name'] == row['Name']:

            #Create email reply text
            email_reply_line = timesheet_key_value['person_name'] + '/' + timesheet_key_value['person_number'] + ' (' + timesheet_key_value['total_time'] + ' timer)\n'
            full_email_string += email_reply_line

    
print(full_email_string)


