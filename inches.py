#!/usr/bin/env python
#-*- coding:utf-8 -*-
# Example from: https://python-docx.readthedocs.io/en/latest/user/text.html

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
document = Document()
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format

print("----paragraph_format.alignment:")
print(paragraph_format.alignment),
paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
print(paragraph_format.alignment)

print("----paragraph_format.alignment:")
