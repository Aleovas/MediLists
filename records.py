#Importing dependencies
import pyautogui
import time
from PIL import Image, ImageEnhance
import pytesseract
import ctypes
import re
import openpyxl
import shutil
import csv
from datetime import date
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, Inches

# Round Order toggle. Set to false to revert to alphabetical room order for list sorting
# When set to true, list order starts with the higher floors and works its way down, with list order for 3B/4B reversed
ROUND_ORDER = True

# This command hides the window for the program when after it starts to not cover the Vista window
ctypes.windll.user32.ShowWindow( ctypes.windll.kernel32.GetConsoleWindow(), 6 )

# Workaround for common OCR errors when reading patient's rooms
def roomClear(x):
    x=x.replace("2-","2A-")
    x=x.replace("28-","2B-")
    x=x.replace("A0","A-0")
    x=x.replace("B0","B-0")
    x=x.replace("C1","C-1")
    x=x.replace("AT","A-1")
    if x[0]=="S": x="3"+x[1:]
    if x[0]=="A": x="4"+x[1:]
    if x[0]=="B": x="5"+x[1:]
    if x[1]=="4": x=x[0]+"A"+x[2:]
    if x[0:2]=="33": x=x.replace("33","3")
    if x[0:2]=="44": x=x.replace("44","3")
    if x[0:2]=="55": x=x.replace("55","3")
    x=x.replace("4AC","4C")
    x=x.replace("5BC","5C")
    if x[2]=="0": x=x[0:2]+"-"+x[3:]
    if x[2]==" ": x=x[0:2]+"-"+x[3:]
    x=x.replace("A1","A-1")
    x=x.replace("B1","B-1")
    if x[-1]=="4": x=x[:-1]+"A"
    if x[-1]=="8": x=x[:-1]+"B"
    x=x.replace("3H","9H")
    x=x.replace("AA","A")
    x=x.replace("-A","-4")
    x=x.replace("-O","-0")
    x=x.replace("-2A","-24")
    x=x.replace("-2B","-28")
    x=x.strip("-")
    x=x.replace("5LD-TX","BLD-TX")
    x=x.replace("DAYCAS","DAYCASE")
    return x

# Workaround for common OCR errors when reading patient's MRN
# Known error: Sometimes the number '5' is read as '6'. This issue is unavoidable with current implementation.
#              Possible fixes require image manipulation which will increase time for list generation significantly
def mrnClear(x):
    x=x.strip().strip("\'\"|/\\").strip("\'\"|/\\‘°").replace(" ","")
    x=x.replace("i","1")
    x=x.replace("o","9")
    x=x.replace("D","5")
    if not x or x.isspace(): return ""
    if x[0]=="0": x="5"+x

    return x

# Given a patient, this function returns the floor the patient is on
def getFloor(x):
    dash=x.room.find("-")
    if dash==-1: dash=2
    return x.room[0:dash]
    
# Given a patient, this function returns the room number the patient is in (without the floor)
def getRoom(x):
    dash=x.room.find("-")
    if dash==-1: dash=2
    return x.room[dash:].strip("-")

# Patient class definition. This class mostly functions as a data container and to facilitate list sorting.
class Patient:
    # Class constructor
    def __init__(self, name, room, mrn):
        self.name = name.replace(",,",",")
        self.room = room
        self.mrn = mrn
    
    # The "less than" comparator definition for this class. Python requires defining at least one comparision method for sorting objects.
    # This function works by defining what the "<" operator returns  (i.e. when patient a (self)< patient b (obj) is true).
    # Two sorting methods are presented based on the ROUND_ORDER value as defined above. Both methods require a special case for floors 10 and 11 
    # as basic string comparison works one character at a time and both will be read as 1* (i.e. less than 2)
    def __lt__(self,obj):
        # Rounding order sorting method
        if (ROUND_ORDER):
            if(self.room[0:2]=="11" and not obj.room[0:2]=="11"): return True
            if(not self.room[0:2]=="11" and obj.room[0:2]=="11"): return False
            if(self.room[0:2]=="10" and not obj.room[0:2]=="10"): return True
            if(not self.room[0:2]=="10" and obj.room[0:2]=="10"): return False
            if(getFloor(self)[0]==getFloor(obj)[0]):
                if(getFloor(self)[1]=="C" and not getFloor(obj)[1]=="C"): return True
                if(not getFloor(self)[1]=="C" and getFloor(obj)[1]=="C"): return False
                if(getFloor(self)[1]=="B" and getFloor(obj)[1]=="A"): return False
                if(getFloor(self)[1]=="A" and getFloor(obj)[1]=="B"): return True
            if(getFloor(self)<getFloor(obj)): return False
            if(getFloor(self)>getFloor(obj)): return True
            if(getFloor(self)==getFloor(obj)):
                if(getFloor(self)[0:2]=="4B"): return getRoom(self)>getRoom(obj)
                if(getFloor(self)[0:2]=="3B"): return getRoom(self)>getRoom(obj)
                else:return getRoom(self)<getRoom(obj)
        # Traditional sorting method
        else:
            if(self.room[0:2]=="11" and not obj.room[0:2]=="11"): return False
            if(not self.room[0:2]=="11" and obj.room[0:2]=="11"): return True
            if(self.room[0:2]=="10" and not obj.room[0:2]=="10"): return False
            if(not self.room[0:2]=="10" and obj.room[0:2]=="10"): return True
            return ((self.room) < (obj.room))
 
# Initialization for pyautogui (the input automation library) an pytesseract (the OCR library) 
pyautogui.FAILSAFE = True
pytesseract.pytesseract.tesseract_cmd=r'C:\Users\Oa.16675\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
cancelTop=pyautogui.locateOnScreen("cancel.png").top
cancelLeft=pyautogui.locateOnScreen("cancel.png").left
try:
    spec=pyautogui.locateOnScreen("spec.png")
    pyautogui.click(spec.left,spec.top)
except:
    print("Failed to find specialties button")
teams={
    "Team 1": ([],"team1.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0}),
    "Team 2": ([],"team2.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0}),
    "Team 3": ([],"team3.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0}),
    "Team 5": ([],"team5.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0}),
    "Lymphoma": ([],"teamLymphoma.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0}),
    "Leukemia": ([],"teamLeukemia.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0}),
    "Palliative": ([],"teampal1.png",{"2A":0,"3A":0,"3B":0,"3C":0,"4A":0,"4B":0,"4C":0,"5C":0,"Towers":0})
}
time.sleep(1)
for team in teams.items():
    found = False
    while not found:
        try:
            print(team[1][1])
            box=pyautogui.locateOnScreen(team[1][1])
            pyautogui.click(box.left,box.top)
            found=True
        except:
            try:
                box=pyautogui.locateOnScreen("scroll.png")
                pyautogui.click(box.left,box.top+20)
                time.sleep(1)
            except:
                pyautogui.alert(text=f'Program failed to scroll automatically to find {team[0]}, please scroll manually until it is visible and click OK', title='Scroll failed', button='OK')
    pyautogui.click(464,134)
    pal=False
    for i in range(40):
        time.sleep(0.8)
        nameImage=pyautogui.screenshot(r'C:\Users\Oa.16675\Downloads\python\1.png', region=(cancelLeft-409,cancelTop+45,392,20))
        roomImage=ImageEnhance.Contrast(pyautogui.screenshot(r'C:\Users\Oa.16675\Downloads\python\3.png', region=(cancelLeft-195,cancelTop+152,74,36))).enhance(20)
        mrnImage=ImageEnhance.Contrast(pyautogui.screenshot(r'C:\Users\Oa.16675\Downloads\python\2.png', region=(cancelLeft,cancelTop+80,68,27))).enhance(20)
        name=pytesseract.image_to_string(nameImage).strip().replace(".",",").replace("_",",").replace("|","").strip("\'\"|/\\").strip("\'\"|/\\")
        if not name or name.isspace(): 
            print("break")
            if pal: break
            if(team[1][1]=="teampal1.png"):
                box=pyautogui.locateOnScreen("teampal2.png")
                pyautogui.click(box.left,box.top)
                time.sleep(1)
                pyautogui.click(464,134)
                pal=True
                continue
            break
        mrn=mrnClear(pytesseract.image_to_string(mrnImage, config="--psm 7"))
        
        roomstr=pytesseract.image_to_string(roomImage, lang="eng", config="--psm 7").strip().strip("\'\"|/\\,.").strip("\'\"|/\\").strip()
        if not mrn or mrn.isspace() or "Test" in name or "Pacs" in name or not roomstr or len(roomstr)<=2:
            pyautogui.press('down')
            continue
        room=roomClear(roomstr)
        team[1][0].append(Patient(name,room,mrn))
        team[1][0].sort()
        pyautogui.press('down')
    print(str(team[0]))

excelName= f'{date.today()}.xlsx'
shutil.copy("template.xlsx",excelName)
wb = openpyxl.load_workbook(excelName)
counts = wb.get_sheet_by_name('Counts')
teamNameCells=tuple(counts["A2":"A8"])
override={}
with open('override.csv') as overrideCSV:
    csv_reader = csv.DictReader(overrideCSV)
    for row in csv_reader:
        override[row["mrn"]]=row["team"]
        
for team in teams.items():
    remove=[]
    for patient in team[1][0]:
        if patient.mrn.replace("6","5") in override.keys(): patient.mrn=patient.mrn.replace("6","5")
        if patient.mrn in override.keys() and not override[patient.mrn]==team[0]:
            if not override[patient.mrn]=="NA": teams[override[patient.mrn]][0].append(patient)
            remove.append(patient)
        
        if patient.room[0:2]=='2B':
            remove.append(patient)
        if patient.room[0:2]=='4H':
            remove.append(patient)
        if patient.room[0:2]=='5H':
            remove.append(patient)
    for patient in remove:
        team[1][0].remove(patient)
        
for team in teams.items():
    team[1][0].sort()
    for patient in team[1][0]:
        pattern=re.compile(r'[2-5][ABC]')
        if(pattern.match(patient.room[0:2])): team[1][2][patient.room[0:2]]+=1
        else: team[1][2]["Towers"]+=1
    print(f'{team[0]}: {team[1][2]}')

    for cell in teamNameCells:
        if cell[0].value == team[0]:
            r=cell[0].row
            c=2
            for count in team[1][2].values():
                counts.cell(row=r,column=c).value=count
                c+=1

    teamSheet=wb.get_sheet_by_name(team[0])
    r=2
    for patient in team[1][0]:
        teamSheet.cell(row=r,column=1).value=str(r-1)
        teamSheet.cell(row=r,column=2).value=patient.name
        teamSheet.cell(row=r,column=3).value=patient.room
        teamSheet.cell(row=r,column=4).value=patient.mrn
        r+=1

wb.save(excelName)
while True:
    try:
        wb.save(excelName)
        break
    except:
        pyautogui.alert(text='Please close Excel file and press OK', title='File write error', button='OK')

document = Document()
style = document.styles['No Spacing']
font = style.font
font.name = 'Calibri'
font.size = Pt(16)
font.bold = True
normalstyle = document.styles['Normal']
normalstyle.font.name = 'Calibri'

for team in teams.items():
    title=document.add_paragraph(f'{team[0]} ({date.today()})')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.style = style
    row_number=len(team[1][0])+1
    row_height=int(Inches(8.5)/row_number)
    table=document.add_table(rows=row_number, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    table.rows[0].cells[3].text="MRN"
    table.rows[0].cells[2].text="Room"
    table.rows[0].cells[1].text="Name"
    i=1
    for patient in team[1][0]:
        table.rows[i].height=row_height
        table.rows[i].cells[3].text=patient.mrn
        table.rows[i].cells[3].width=0
        table.rows[i].cells[2].text=patient.room
        table.rows[i].cells[2].width=0
        table.rows[i].cells[1].text=patient.name
        table.rows[i].cells[1].width=Cm(50)
        table.rows[i].cells[0].text=str(i)
        table.rows[i].cells[0].width=0
        i+=1
    if not team[0] == "Palliative": document.add_page_break()   

while True:
    try:
        document.save(f"{date.today()}.docx")
        break
    except:
        pyautogui.alert(text='Please close Word file and press OK', title='File write error', button='OK')    
