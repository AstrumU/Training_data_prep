#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 26 10:41:07 2019

@author: cheng
"""

from bs4 import BeautifulSoup
import requests
import wget

start_url = "https://mcckc.edu/degreescourses/CIF.html"
response = requests.get(start_url)
page = response.content
soup = BeautifulSoup(page, 'lxml')

for link in soup.find_all("a"):
    if 'href' in link.attrs:
        if link['href'].endswith('.docx') or link['href'].endswith('.doc') or link['href'].endswith('.pdf'):
            url = link['href']
            print(url, url.split('/')[-1])
            wget.download(url, '/Users/cheng/Downloads/courses/' + url.split('/')[-1])