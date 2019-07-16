#! python3
# BatchEmailMerge.py - Create draft email with word template according to the information in excel file 
# (for example : Email, Name, attachment, and field)

import os.path, shutil, mammoth, re
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
word = win32.DispatchEx("Word.Application")
outlook = win32.Dispatch("Outlook.Application")
word.Visible = False
excel.Visible = False

# Property for our hotel
Property = ['VMRH', 'PARIS', 'CMCC', 'HICC']

# Find the folder path under RunFile
path1 = os.getcwd() + '\\RunFile\\'
for file in os.listdir(path1):
    Filename = file
path1 = path1 + Filename + '\\'

# If DraftEmail folder exisis, deelete it
if os.path.exists(path1 + 'DraftEmail'):
    shutil.rmtree(path1 + 'DraftEmail')
# Create a new DraftEmail folder
os.mkdir(path1 + 'DraftEmail')

# Open the Excel file and worksheet
# TODO: Change excel filename and tab name
EmailList = path1 + '\\EmailList.xlsx'
wb1 = excel.Workbooks.Open(EmailList)
ws1 = wb1.Worksheets('EmailList')

# List for column number for Field and Attachment
FieldList = []
AttachList = []
# Initital value for y, Field number and Attach number for loop function
y = 1
FieldNum = 1
AttachNum = 1
# To find the column number for Field and Attachment and add to the FieldList and AttachList
while True:
    # If the column header is empty, exit the loop
    if str(ws1.Cells(3, y).Value) == 'None':
        break
    # Find the column number for Field to replace
    if ws1.Cells(3, y).Value == 'Field' + str(FieldNum):
        FieldList.append(y)
        FieldNum += 1
    # Find the column number for Attachment
    if ws1.Cells(3, y).Value == 'Attach' + str(AttachNum):
        AttachList.append(y)
        AttachNum += 1
    y += 1

# Email draft need to create if Batch name match
BatchToRun = ws1.Cells(1, 2).Value
# Word file template used to create email draft
WordFilename = ws1.Cells(2, 2).Value
# Dictionary with path for email and number of email drafted
AssistantList = {}
# x is to find the End of the row number
x = 4
while True:
    # If the Assistant field is empty, exit the loop
    if str(ws1.Cells(x, 2).Value) == 'None':
        break
    # If Batch field is empty will pass the loop
    if ws1.Cells(x, 1).Value == BatchToRun:
        Name = ws1.Cells(x, 2).Value
        # Create Assistant Folder if it does not exist
        if Name not in AssistantList:
            path2 = path1 + '\\DraftEmail\\'
            os.mkdir(path2 + Name)
            AssistantList.setdefault(Name, []).append(str(path2 + Name))
            AssistantList[Name].append(1)
        else:
            AssistantList[Name][1] = AssistantList[Name][1] + 1
            
        # Open word template to replace
        doc = word.Documents.Open(path1 + WordFilename, False, False)
        Num = 1
        for field in FieldList:
            word.Selection.Find.Execute(str('[Field' + str(Num) + ']'), False, False, False, False, False, True, 1, True, str(ws1.Cells(x, field).Value), 2)
            Num += 1
        
        # Sentence for Properties in email body
        if str(ws1.Cells(x, 3).Value) != 'None':
            PropertyInclude = str(ws1.Cells(x, 3).Value).split()
            # Compare the difference 
            DelProperty = list(set(Property) - set(PropertyInclude))
            # re function for look up
            PropertyDel = re.compile(r'\[%s\]' % '|'.join(DelProperty))
            
            # ParagraphIndex to save the paragraph row number for delete purpose
            ParagraphIndex = []
            for parNo in range(1, doc.Paragraphs.Count):
                current_text = doc.Paragraphs(parNo).Range.Text
                # Use re function to search in each paragraph
                PropertySearch = PropertyDel.search(current_text)
                if PropertySearch != None:
                    ParagraphIndex.append(parNo)
            
            # Sort by descending (to delete paragraph from bottom first)
            ParagraphIndex.sort(reverse=True)
            # Delete Paragrapth by index number according to ParagraphIndex list
            for no in ParagraphIndex:
                doc.Paragraphs(no).Range.Delete()
            # TODO: Remove the remaining property name in Word
            
        # Amend the file name for Template file
        doc.SaveAs(path1 + '\\Template.docx', FileFormat=12)
        doc.Close()
        
        # Convert the Word file into HTML text
        with open(path1 + '\\Template.docx', 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        
        # Search for keyWord [Image1] in html to creat Image List
        keyWord1 = re.compile(r'(\[Image\d\])')
        ImageList = keyWord1.findall(html)
        # Search for keyWord [ImageRegion1] in html to creat Image List
        keyWordCountry = re.compile(r'(\[ImageRegion\d\])')
        CountryImageList = keyWordCountry.findall(html)
        
        # Create draft email in outlook
        mail = outlook.CreateItem(0)
        mail.To = str(ws1.Cells(x, 6).Value)
        mail.CC = str(ws1.Cells(x, 7).Value)
        mail.BCC = str(ws1.Cells(x, 8).Value)
        mail.Subject = str(ws1.Cells(x, 5).Value)
        
        # Add Image to html body
        ImgNum = 1
        if len(ImageList) > 0:
            for image in ImageList:
                # Use keyWord to replace the [Image] in HTML text
                keyWord2 = re.compile(r'(\[%s\])' % image[1:-1])
                html = keyWord2.sub("<img src=""cid:MyId%s"">" % ImgNum, html)
                # Change Image size if needed
                #html = keyWord2.sub("<a href=""mailto:walter.loo@sands.com.mo""><img src=""cid:MyId%s"" align=""middle"" height=""1700"" width=""1000""></a>" % ImgNum, html)
                attachment = mail.Attachments.Add(path1 + "\\Image\\" + image[1:-1] + ".jpg", 0x5, 0)
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId%s" % ImgNum)
                ImgNum += 1
        """
        # Add Image for Country to html body
        if len(CountryImageList) > 0:
            CountryImgNum = 1
            for image in CountryImageList:
                html = keyWordCountry.sub("<img src=""cid:MyId%s"">" % ImgNum, html)
                attachment = mail.Attachments.Add("I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\BatchEmailMerge\\Image\\Image%s%s.jpg" % (str(ws1.Cells(x, 4).Value), str(CountryImgNum)), 0x5, 0)
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId%s" % ImgNum)
                CountryImgNum += 1
                ImgNum += 1
        """
        mail.HtmlBody = "<html><body>" + html +  "</body></html>"
        # Add Attachment to the email
        for field in AttachList:
            mail.Attachments.Add(path1 + '\\Attachment\\'  + str(ws1.Cells(x, field).Value))
            
        # SaveAs the file in the Assistant folder
        mail.SaveAs(Path=AssistantList[Name][0] + '\\' + ws1.Cells(x,2).Value + str(AssistantList[Name][1]) + '.msg')
        
    x += 1

# Delete word file created after running the script
os.unlink(path1 + '\\Template.docx')
