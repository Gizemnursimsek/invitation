# -*- coding: utf-8 -*-
"""
Created on Fri Jul  8 16:48:54 2022

@author: W10
"""

import os
import pandas as pd
from docx import Document

os.chdir ("C:/Users/User/Desktop/python codes")
df = pd.read_excel("guests.xlsx", sheet_name = "page1")

names = df["guests"]

def change (words):
    file = Document("invitation example.docx")
    for paragraph in dosya.paragraphs:
        if "X" in paragraph.text:
            paragraph.text = "Dear " + word
            dosya.save(word + ".docx")


for i in names:
    change(i)
