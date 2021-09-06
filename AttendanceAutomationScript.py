""" code for sorting attendance data taken from Microsoft Teams """
#!/usr/bin/env python3
import csv
import sys 
from datetime import date
from collections import OrderedDict
import codecs
import smtplib
import PySimpleGUI as sg


def file_path_gui():
    """ This Function creates a gui for selecting a file and outputing the file Path"""
    sg.theme("DarkTeal2")
    layout = [[sg.T("")],
              [sg.Text("Choose the attendance file: "),
               sg.Input(),
               sg.FileBrowse(key="-IN-")],
              [sg.Button("Submit")]]
    window = sg.Window('Attendance_Report', layout, size=(600, 150))

    while True:
        event, values = window.read()
        if event in(sg.WIN_CLOSED, 'Exit'):   #event == sg.WIN_CLOSED or event == "Exit":
            sys.exit()
        elif event == "Submit":
            print(values["-IN-"])
            window.close()
            fpath = values["-IN-"]
        return fpath

def user_data():
    """ This function gets user data (name, email and class name) where email is required and is used for sending mail"""
    sg.theme("DarkTeal2")

    layout = [
    [sg.Text("Enter User Details '*' Required Fields" )],
    [sg.Text('Name', size = (15, 1)), sg.InputText()],
    [sg.Text('Email ID*', size = (15, 1)), sg.InputText()],
    [sg.Text('Class Name', size = (15, 1)), sg.InputText()],
    [sg.Submit(), sg.Cancel()]]
    window = sg.Window('Enter User Details :)', layout)
    event, values = window.read()
    window.close()

    global class_name
    global user_name
    global email_id
    user_name = values[0]
    email_id = values[1]
    class_name = values[2]

    if not email_id:
        sg.popup("NO EMAIL_ID GIVEN ! The report will not be sent through mail :)")
    
def csv_at_data(file_path):
    """ this function takes reads the csv document and filters the attendance data and
    stores it in a list"""
    column_name = "Full Name"
    read_encoding_format = "utf-8"
    if '\0' in open(file_path).read():
        read_encoding_format = 'utf-16'

    with codecs.open(file_path, 'rU', read_encoding_format) as file:
        reader = csv.DictReader(file, delimiter='\t')
        raw_names_list = []

        for row in reader:
            names = row[column_name]
            raw_names_list.append(names)
            new_list = list(OrderedDict.fromkeys(raw_names_list))
    return new_list

user_data()

try:
    at_list = csv_at_data(file_path_gui())   
except Exception as error:
    sg.popup("unable to process attendance report check file type is a csv document:)", error)
    sys.exit()

try:
    today = date.today()
except Exception as error:
    sg.popup("error getting date and time :)", error)

JOIN_DATA = '\n'.join([str(i) for i in at_list])
sg.PopupScrolled("Attendance Report","class",class_name,
                 "date = ", "click ok to send data via mail :)", today, f"{JOIN_DATA}")

try:
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()

        smtp.login('enter mail id here', 'enter password here')

        subject = f'Attendance report of: {today}'
        body = JOIN_DATA
        msg = f'subject: {subject}\n\nUser Name: {user_name}\n\nclass Name: {class_name}\n\n{body}'
        smtp.sendmail('enter mail id here', email_id, msg)
        smtp.quit()
except Exception as error:
    sg.popup("Unable to send attendance report via Mail please check Mail ID :)", error)
    sys.exit()

sg.popup("Attendance Report Has been sent to your mail id :) ")
