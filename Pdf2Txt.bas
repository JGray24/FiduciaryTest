Attribute VB_Name = "Pdf2Txt"
Option Compare Database
Option Explicit

'Found this on the below link.
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=977

' 1) First of all download the freeware pdf utilities at http://www.foolabs.com/xpdf/download.html.

' 2) Although there are several programs in the download the only one you need to keep is pdftotext.exe.
'    This utility will extract text from pdf files, so scanned pdf documents (containing only images)
'    will result in an error.

' 3) Copy pdftotext.exe in a folder called C:\pdf2txt.

' 4) Open Notepad and type the following: pdftotext.exe -layout YourPage.pdf

' 5) Save the bat-file containing the above line in the pdf2txt folder as Yourpage.bat
'    (not YourPage.bat.txt or YourPage.txt).

'Examples:

'pdftotext -f 1 -l 1 C:\Users\johnr\Desktop\XpdfTest\UnionState.PDF C:\Users\johnr\Desktop\XpdfTest\UnionState.txt

'pdftotext -f 1 -l 1 "C:\Users\johnr\Desktop\XpdfTest\Union State.PDF" "C:\Users\johnr\Desktop\XpdfTest\UnionStatexx.txt"

'pdftotext -f 1 -l 2 "C:\Users\johnr\Desktop\VA Services Design\VA Services Data Repository\Veterans\Brown, Terry C423809994\Checking\Terry Brown  20170500 Statements_07182017_215606.PDF" "C:\Users\johnr\Desktop\XpdfTest\UnionState xxx.txt"

'pdftotext -f 1 -l 2 "C:\Users\johnr\Desktop\VA Services Design\VA Services Data Repository\Veterans\Brown, Joe P  C416786803\Checking\170989 1 C Sep+9%2C+2017.PDF" "C:\Users\johnr\Desktop\XpdfTest\Joe P Brown.txt"

'pdftotext -f 1 -l 2 "C:\Users\johnr\Desktop\XpdfTest\170989 1 C Sep+9%2C+2017_recognized_1.pdf" "C:\Users\johnr\Desktop\XpdfTest\Joe P Brown OCRd.txt"

'pdftotext -f 1 -l 2 "C:\Users\johnr\Desktop\XpdfTest\UnionState JGray_recognized_1.pdf" "C:\Users\johnr\Desktop\XpdfTest\UnionState JGray.txt"




Sub OpenPDF()
     '-------------------------------------------'
     ' You need to create a bat file first with one single line of text
     ' pdftotext.exe -layout YourPage.pdf
     ' DOWNLOAD LINK: http://www.foolabs.com/xpdf/download.html
     '-------------------------------------------'
     
     ' these lines look for a pdf file in your My Documents folder
    Set WshShell = CreateObject("WScript.Shell")
    ChDir (WshShell.SpecialFolders("MyDocuments"))
    PageName = Application.GetOpenFilename("YourPage, *.pdf", , "YourPage")
     
     ' if no file is picked the macro ends
    If PageName = "False" Then
        Exit Sub
    End If
     
     ' copies and renames the pdf file to the pdf2txt folder
    FileCopy PageName, "C:\pdf2txt\YourPage.pdf"
    ChDir ("C:\pdf2txt")
     
     ' THE BATFILE CONTAINS ONLY 1 LINE:
     ' pdftotext.exe -layout YourPage.pdf
     
    TestValue = Shell("YourPage.bat", 1)
     
     ' because the bat file runs for 1 or 2 seconds (in my case)
     ' I let the Excel macro wait 5 seconds before doing anything else
     ' there are more ingenious ways for VBA to wait for the end of an
     ' application, but this suits me fine...
     
    Application.Wait (Now + TimeValue("0:00:05"))
     
    ChDir "C:\pdf2txt"
    PageName = "C:\pdf2txt\YourPage.txt"
     
     ' the following reads the text that has been generated
    Call ReadTextFile
     
     ' insert your text parsing - text to columns - ingenious vba stuff hereafter...
     
End Sub
 
Sub ReadTextFile()
    Dim FileNum As Long
    Dim r As Long
    Dim wb As Workbook
    Dim Data As String
    r = 1
    FileNum = FreeFile
    Set wb = Workbooks.Add
    Open PageName For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, Data
        ActiveSheet.Cells(r, 1) = Data
        r = r + 1
    Loop
    Close #FileNum
End Sub

