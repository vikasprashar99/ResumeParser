
# NECCESSARY IMPORTS
from django.core.files.storage import FileSystemStorage
from django.http import HttpResponse
from django.shortcuts import render

# pipn install PyMuPDF
import fitz
# pip install docx
import docx
import re
import zipfile
# pip i docx2txt
import docx2txt
# pip i pymongo
import pymongo
#pip i xwlt
import xlwt
#pip i dns

# connecting with the mongoDB database
client = pymongo.MongoClient(
    "mongodb+srv://vikas:Test123@node-events-j0bd8.mongodb.net/test?retryWrites=true&w=majority")
db = client.get_database('djangoDb')
records = db.Users

# FUNCTION THAT WILL BE CALLED FIRST WHEN PAGE LOADS
def dashboard(request):
    return render(request, "resumeUpload.html")

# WHEN THE USER WILL UPLOAD FILE
def getResumeData(request):
    f = request.FILES["filename"]
    fs = FileSystemStorage()
    filename = fs.save(f.name, f)

    global newCreatedfilename
    newCreatedfilename=f.name.split(".")[0]
    fileExtension = f.name.split(".")[1]

    
    #**************************************************FOR WORD OR DOCX FILE****************************************************************** 
    if fileExtension == "docx" or fileExtension == "doc":
        wordText = docx2txt.process(f)
        noOfCharacters=len(wordText)

        # FOR GETTING THE NAME
        firstText=wordText.split("\n")
        if firstText[0]=="RESUME" or firstText[0]=="Resume":
            firstText=firstText[2]+""+firstText[3]
            Name=firstText
        else:
            Name = firstText[0]
        print(firstText[0])

        # FOR GETTING THE EMAIL
        emailPattern = re.compile(r'[a-zA-Z0-9-\.]+@[a-zA-Z-\.]*\.(com|edu|net)')
        EmailmatchesDoc = emailPattern.finditer(wordText)
        Email = "None"
        for match in EmailmatchesDoc:
            Email = match.group(0)

        # FOR GETTING THE LINKEDIN PROFILE
        linkedinPattern = re.compile('^https:\\/\\/[a-z]{2,3}\\.linkedin\\.com\\/.*$')
        linkedINmatches=linkedinPattern.finditer(wordText)
        linkedIN="None"
        for match in linkedINmatches:
            linkedIN=match.group(0)

        # FOR GETTING THE PHONE NUMBER
        phonePatternDoc = re.compile(r'(\(|\+)?\d{3}(\)|\-)?\s?\d{2,3}\-?\d{3,6}')
        phoneMatches = phonePatternDoc.finditer(wordText)
        Phone_Num = 0
        for match in phoneMatches:
            Phone_Num = match.group(0)

        # FOR GETTING TABLES
        document = docx.Document(f)
        Table_Count = 0
        Table_Count = len(document.tables)

        # FOR FONT NAME AND SIZE
        doc = docx.Document(f)        
        fontSize=[]
        for p in doc.paragraphs:
            name = p.style.font.name
            size = p.style.font.size
            fontSize.append(size)
        font_name=name
        font_size=fontSize[0]

        # FOR TOTAL IMAGES
        totalImages = []
        z = zipfile.ZipFile(f)
        all_files = z.namelist()
        images = filter(lambda x: x.startswith('word/media/'), all_files)
        Images = 0
        for match in images:
            totalImages.append(match)
            Images = len(totalImages)
 #********************************* FOR PDF FILE*********************************************************************************************
    else:
        file1 = fitz.open(f)
        pdfText = file1.getPageText(0)

        # FOR NUMBER OF CHARACTERS
        noOfCharacters=len(pdfText)

        # FOR GETTING THE NAME
        firstText=pdfText.split(" ")
        if firstText[0]=="RESUME" or firstText[0]=="Resume":
            firstText=firstText[2]+firstText[3]
            Name=firstText
        else:
            Name = pdfText.split("  ")[0].split("\n")[0]

         # FOR GETTING THE EMAIL
        emailPattern = re.compile(
            r'[a-zA-Z0-9-\.]+@[a-zA-Z-\.]*\.(com|edu|net)')
        Emailmatches = emailPattern.finditer(pdfText)
        Email = "None"
        for match in Emailmatches:
            Email = match.group(0)

        # FOR GETTING THE LINKEDIN PROFILE
        linkedinPattern = re.compile('^https:\\/\\/[a-z]{2,3}\\.linkedin\\.com\\/.*$')
        linkedINmatches=linkedinPattern.finditer(pdfText)
        linkedIN="None"
        for match in linkedINmatches:
            linkedIN=match.group(0)

        # FOR GETTING THE PHONE NUMBER
        phonePattern = re.compile(r'(\(|\+)?\d{3}(\)|\-)?\s?\d{2,3}\-?\d{3,6}')
        phoneMatches = phonePattern.finditer(pdfText)
        Phone_Num = 0
        for match in phoneMatches:
            Phone_Num = match.group(0)
        if(Phone_Num == 0):
            Phone_Num = "None"

      # FOR TOTAL IMAGES
        Images = len(file1.getPageImageList(0))
        Table_Count = 0

         # FOR FONT NAME AND SIZE
        for font in range(len(file1.getPageFontList(0))):
            fontname=file1.getPageFontList(0)[font][3]
            fontsize=file1.getPageFontList(0)[font][0]
        font_name=fontname
        font_size=fontsize

    # STORING THE DATA IN PARAMS TO BE SHOWN ON NEXT PAGE AND SENDIN THE SAME TO DATABSE
    params = {"Name": Name, "Email": Email, "Phone": Phone_Num, "Images": Images,"LinkedIN":linkedIN , "Font_name":font_name,
    "Font_size":font_size,"Tables":Table_Count,"Noofcharacters":noOfCharacters}
    records.insert_one(params)

    # ROUTING TO NEW PAGE
    return render(request, "resumeData.html", params)


def upload_func(file):
    with open('../DjangoApp/mysite/filename.pdf', 'wb+') as f:
        for chunk in file.chunks():
            f.write(chunk)

# FUNC TO DOWNLOAD EXCEL FILE BUTTON
def downloadCSV(request):

    # GETTING THE RECORDS FROM THE DB
    data=list(records.find())

    # CREATING A WORKBOOK(EXCEL)
    wb=xlwt.Workbook()
    ws=wb.add_sheet("new sheet")
    ws.write(0,0,"Name")
    ws.write(0,1,"Email")
    ws.write(0,2,"Phone")
    ws.write(0,3,"LinkedIN")
    ws.write(0,4,"Tables")
    ws.write(0,5,"Images")
    ws.write(0,6,"Font_size")
    ws.write(0,7,"Font_name")
    ws.write(0,8,"Total number of characters")

    for i in range(len(data)):
        ws.write(i+1,0,data[i]["Name"])
        ws.write(i+1,1,data[i]["Email"])
        ws.write(i+1,2,data[i]["Phone"])
        ws.write(i+1,3,data[i]["LinkedIN"])
        ws.write(i+1,4,data[i]["Images"])
        ws.write(i+1,5,data[i]["Tables"])
        ws.write(i+1,6,data[i]["Font_name"])
        ws.write(i+1,7,data[i]["Font_size"])
        ws.write(i+1,8,data[i]["Noofcharacters"])

    wb.save(newCreatedfilename+".xls")

    return HttpResponse("YOUR FILR HAS BEEN DOWNLOADED BY THE SAME NAME YOU HAVE UPLOADED IN THE SAME DIRECTORY OF THE PROJECT")
