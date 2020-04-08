# created this file
from django.http import HttpResponse
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
import fitz
# pip install docx
import docx
import re
import zipfile
# pip i docx2txt
import docx2txt
# pip i pymongo
import pymongo

import xlwt

# let Testfilename

client = pymongo.MongoClient(
    "mongodb+srv://vikas:Test123@node-events-j0bd8.mongodb.net/test?retryWrites=true&w=majority")
db = client.get_database('djangoDb')
records = db.Users


def dashboard(request):
    return render(request, "index.html")


def getImage(request):
    f = request.FILES["filename"]
    fs = FileSystemStorage()
    filename = fs.save(f.name, f)
    uploaded_file_url = fs.url(filename)
    global Testfilename
    Testfilename=f.name.split(".")[0]

    ext = f.name.split(".")[1]
    if ext == "docx" or ext == "doc":
        my_text = docx2txt.process(f)
        x2 = my_text.split("  ")[0].split("\n")[0]
        emailPattern = re.compile(
            r'[a-zA-Z0-9-\.]+@[a-zA-Z-\.]*\.(com|edu|net)')
        EmailmatchesDoc = emailPattern.finditer(my_text)
        x4 = "None"
        for match in EmailmatchesDoc:
            x4 = match.group(0)

        phonePatternDoc = re.compile(
            r'(\(|\+)?\d{3}(\)|\-)?\s?\d{2,3}\-?\d{3,6}')
        phoneMatches = phonePatternDoc.finditer(my_text)
        x5 = "none"
        for match in phoneMatches:
            x5 = match.group(0)

        # FOR TABLES
        document = docx.Document(f)
        table_count = "none"
        for table in document.tables:
            table_count = len(document.tables)
        print(table_count)

        doc = docx.Document(f)

        for p in doc.paragraphs:
            name = p.style.font.name
            size = p.style.font.size
            # print(name,size)

        totalImages = []
        z = zipfile.ZipFile(f)
        all_files = z.namelist()
        images = filter(lambda x: x.startswith('word/media/'), all_files)
        x6 = "0"
        for match in images:
            totalImages.append(match)
            x6 = len(totalImages)
    else:
        file1 = fitz.open(f)
        pages = len(file1)
        x1 = file1.getPageText(0).split("\n ")[1]
        pdfText = file1.getPageText(0)
        x2 = pdfText.split("  ")[0].split("\n")[0]
        firstText = pdfText.split("  ")[0].split("\n")[0]
        emailPattern = re.compile(
            r'[a-zA-Z0-9-\.]+@[a-zA-Z-\.]*\.(com|edu|net)')
        Emailmatches = emailPattern.finditer(pdfText)
        x4 = 0
        for match in Emailmatches:
            x4 = match.group(0)

        phonePattern = re.compile(r'(\(|\+)?\d{3}(\)|\-)?\s?\d{2,3}\-?\d{3,6}')
        phoneMatches = phonePattern.finditer(pdfText)
        x5 = 0
        for match in phoneMatches:
            x5 = match.group(0)
        if(x4 == 0):
            x4 = "None"
        if(x5 == 0):
            x5 = "None"
        x6 = len(file1.getPageImageList(0))

    
    params = {"name": x2, "email": x4, "phone": x5, "lenimage": x6}
    
   
    records.insert_one(params)
    print(list(records.find()))

    

    return render(request, "nextpage.html", params)


def upload_func(file):
    with open('../DjangoApp/mysite/filename.pdf', 'wb+') as f:
        for chunk in file.chunks():
            f.write(chunk)

def test(request):
    print(list(records.find()))
    data=list(records.find())
    wb=xlwt.Workbook()
    wb.set_tab_width(100)
    ws=wb.add_sheet("new sheet")

    ws.write(0,0,"Name")
    ws.write(0,1,"Email")
    ws.write(0,2,"Phone")
    ws.write(0,3,"Images")

    for i in range(len(data)):
        ws.write(i+1,0,data[i]["name"])
        ws.write(i+2,1,data[i]["email"])
        ws.write(i+3,2,data[i]["phone"])
        ws.write(i+4,3,data[i]["lenimage"])

    wb.save(Testfilename+".xls")

    return HttpResponse("Successfully got it")
