#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Download Course Syllabus in PDF for Clemson University

@author: cheng
"""
from bs4 import BeautifulSoup
import requests
import wget
import pandas as pd

READ_EXCEL_FILE = '/Users/cheng/Desktop/Clemson_University/Discipline.csv'
DF = pd.read_csv(READ_EXCEL_FILE)
COURSES = []

URL = 'https://syllabus.app.clemson.edu/repository/syllabus_public.php'
R1 = requests.post(URL, data={"person_ck": "CLEMSON", "check_person": "Logon"})

for i in range(DF.iloc[:, 0].size):
    buff = DF.loc[i][0]
    COURSES.append(buff)

for course in COURSES:
    session_id = R1.cookies.get('PHPSESSID')
    R2 = requests.post(URL, data={"semester_selected": "spring:2016", "subj_selected": course,
                                  "search_course": "Load Course Files"},
                       cookies={'PHPSESSID': session_id})

    soup = BeautifulSoup(R2.content, 'lxml')

    for link in soup.find_all("a"):
        if 'href' in link.attrs:
            if link['href'].endswith('spring2016.pdf'):
                newurl = 'https://syllabus.app.clemson.edu' + link['href'][2:]
                middle = link.contents[0][:link.contents[0].find('.')]
                middle = middle.replace(' ', '_')
                wget.download(newurl, '/Users/li/Downloads/' + middle +'_'+ newurl.split('/')[-1])
