#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 26 10:41:07 2019
Download Course Syllabus in doc or pdf for MCC
@author: cheng
"""
from bs4 import BeautifulSoup
import requests
import wget

START_URL = "https://mcckc.edu/degreescourses/CIF.html"
RESPONSE = requests.get(START_URL)
PAGE = RESPONSE.content
SOUP = BeautifulSoup(PAGE, 'lxml')

for link in SOUP.find_all("a"):
    if 'href' in link.attrs:
        if (link['href'].endswith('.docx') or link['href'].endswith('.doc')
                or link['href'].endswith('.pdf')):
            url = link['href']
            print(url, url.split('/')[-1])
            wget.download(url, '/Users/cheng/Downloads/courses/' + url.split('/')[-1])
