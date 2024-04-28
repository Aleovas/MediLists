# MediLists
This is a simple python script to allow for automatic list generations of CPRS electronic health records systems. This program uses PyAutoGui for screenshots and GUI automation and pytesseract for OCR. 

# Current implemented features
- Automatically finding team lists and patients
- Creating excel sheets with team counts and lists
- Comparision with previous excel file, if present, to allow for statistics regarding admissions and discharges/transfers
- Creating word documents with lists for easier printing

# Planned features
- A proper GUI for user-friendliness
- More configuration options

# Known bugs
- Currently only supports a font size of 12, this is being worked on

# Installation instructions
- Install [python](https://www.python.org/downloads/) and [tesseract](https://github.com/UB-Mannheim/tesseract/wiki) 
- Install requirements from requirements.txt using pip
- Run records.py
