#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Download Course Syllabus in PDF for Clemson University

@author: cheng
"""
from bs4 import BeautifulSoup
import requests
import wget

URL = 'https://syllabus.app.clemson.edu/repository/syllabus_public.php'
R1 = requests.post(URL, data={"person_ck": "CLEMSON", "check_person": "Logon"})

TEMP = 'AAH, ACCT, AGED, AGM, AGRB, AL, ANTH, APEC, ARAB, ARCH, ART, AS, ASL, ASTR, AUD, AUE, AVS,\
        BCHM, BE, BIOE, BIOL, BMOL, BUS, CE, CES, CH, CHE, CHIN, COMM, COOP, CPSC, CRP, CSM, CU,\
        CVT, DANC, DPA, ECE, ECON, ED, EDC, EDEC, EDEL, EDF, EDL, EDLT, EDML, EDSC, EDSP, EES, ELE,\
        ENGL, ENGR, ENR, ENSP, ENT, ESED, ETOX, FCS, FDSC, FDTH, FIN, FNR, FOR, FR, GC, GEN, GEOG,\
        GEOL, GER, GW, HCC, HCG, HEHD, HIST, HLTH, HON, HORT, HP, HRD, HUM, IE, INT, IPM, ITAL,\
        JAPN, LANG, LARC, LAW, LIB, LIH, LIT, LS, MATH, MBA, ME, MGT, MHA, MICR, MKT, ML, MSE, MUSC,\
        NPL, NURS, NUTR, PA, PADM, PAS, PDBE, PES, PHIL, PHSC, PHYS, PKSC, PLPA, POSC, POST, PRTM,\
        PSYC, RCID, RED, REL, RS, RUSS, SOC, SPAN, STAT, STS, THEA, WFB, WS, YDP'

TEMP = TEMP.replace(' ', '')
COURSES = TEMP.split(',')

for course in COURSES:
    session_id = R1.cookies.get('PHPSESSID')
    R2 = requests.post(URL, data={"semester_selected": "summer:2015", "subj_selected": course,
                                  "search_course": "Load Course Files"},
                       cookies={'PHPSESSID': session_id})

    soup = BeautifulSoup(R2.content, 'lxml')

    for link in soup.find_all("a"):
        if 'href' in link.attrs:
            if link['href'].endswith('summer2015.pdf'):
                newurl = 'https://syllabus.app.clemson.edu' + link['href'][2:]
                middle = link.contents[0][:link.contents[0].find('.')]
                print(middle)
                middle = middle.replace(' ', '_')
                wget.download(newurl, '/courses/' + middle +'_'+ newurl.split('/')[-1])
