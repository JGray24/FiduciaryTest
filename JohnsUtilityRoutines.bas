Attribute VB_Name = "JohnsUtilityRoutines"
Option Compare Database
Option Explicit

' DelTbl              Delete a table from the CurrentDB
' Scrub               Ensure that all apostrophe characters are translated out before passing to SQL
' Remove_Table_Field  Remove a field from a table definition
' FieldExists         Test if fieldName field exists in tableName table
' Create_A_Table      Create a table with an AutoNumber Id primary key so that new fields can be added.  (Delete existing table first)

Dim ErrorMsg   As String

Public Sub DelTbl(TblName As String)
Dim errMsg    As String
Dim Response  '  Yes=6  No=7  Retry=4  OK=1  Cancel=2
   
On Error GoTo Error_Handler
   
Application.CurrentDb.TableDefs.Delete TblName
On Error GoTo 0
Exit Sub
   
Error_Handler:
  Select Case err.Number
  Case 3211
    errMsg = "Error number: " & Str(err.Number) & vbNewLine & _
             "Source: " & err.source & vbNewLine & _
             "Description: " & err.Description
    Response = MsgBox(errMsg, vbRetryCancel)
    Debug.Print (Chr(13) & "****" & errMsg & Chr(13))
    If Response = 4 Then Resume
    End
  Case Else
    Resume Next
  End Select
  
End Sub
Public Function DelTblS(TblName As String)

' This routine is similar to "DelTbl", but has support for WildCard "*" and can delete multiple tables.

Dim db As dao.Database
Dim tdf As dao.TableDef
Set db = CurrentDb
For Each tdf In db.TableDefs
    ' ignore system and temporary tables
    If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") And (tdf.Name Like TblName) Then
        Debug.Print ("Deleting DB Table - " & tdf.Name)
        Call DelTbl(tdf.Name)
    End If
Next
Set tdf = Nothing
Set db = Nothing

End Function


Public Function Scrub(x As String)
'  Ensure that all apostrophe characters are translated out before passing to SQL
Dim I   As Long, ret As String, aChar  As String, aChr As String
For I = 1 To Len(x)
  aChar = Mid(x, I, 1)
  aChr = ""
  If aChar = "'" Then aChr = "chr(39)"
 ' If aChar = ":" Then aChr = "chr(158)"
  If aChr <> "" Then
    If Len(ret) = 0 Then
      ret = aChr & " & '"
    Else
    If Mid(ret, Len(ret), 1) = "'" Then
      ret = Left(ret, Len(ret) - 1) & aChr & " & '"
    Else
    ret = ret & "' & " & aChr & " & '"
    End If
    End If
  End If
  If aChr = "" Then
    If ret = "" Then
      ret = "'"
    End If
    ret = ret & aChar
  End If
Next I
If Right(ret, 4) = " & '" Then
  ret = Left(ret, Len(ret) - 4)
Else
  ret = ret & "'"
End If

Scrub = ret
End Function

Public Sub Remove_Table_Field(fieldName As String, tableName As String)

Dim tdf       As TableDef
Dim db As dao.Database
Dim fld As dao.Field
Dim prop As dao.Property

If Not FieldExists(fieldName, tableName) Then Exit Sub

Set db = CurrentDb
Set tdf = db.TableDefs(tableName)

tdf.Fields.Delete (fieldName)

Set tdf = Nothing
Set tdf = Nothing

End Sub

' test if fieldName field exists in tableName table
Public Function FieldExists(fieldName As String, tableName As String) As Boolean
    Dim db As dao.Database
    Dim tbl As TableDef
    Dim fld As Field
    Dim strName As String
    Set db = CurrentDb
    Set tbl = db.TableDefs(tableName)
    FieldExists = False
    For Each fld In tbl.Fields
        If fld.Name = fieldName Then
            FieldExists = True
            Exit For
        End If
    Next
End Function


Public Sub Create_A_Table(theTableName As String)
 
' Create a table with an AutoNumber Id primary key so that new fields can be added.  (Delete existing table first)
Dim strSql As String
DelTbl (theTableName)
strSql = "CREATE TABLE " & theTableName & " (id COUNTER PRIMARY KEY);"
'Debug.Print (strSql)
DoCmd.RunSQL (strSql)
 
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RunFile
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Run/Execute files from vba (bat, vbs,…)
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - full path including filename and extension
' strWndStyle - style of the window in which the program is to be run
'               value can be vbHide,vbNormalFocus,vbMinimizedFocus
'               vbMaximizedFocus,vbNormalNoFocus or vbMinimizedNoFocus
' Usage Example:
' ~~~~~~~~~~~~~~~~
' RunFile("c:\test.bat", vbNormalFocus)
' Revision History:
' ~~~~~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2010-Feb-05             Initial Release
'---------------------------------------------------------------------------------------
Function RunFile(strFile As String, strWndStyle As String)
On Error GoTo Error_Handler
 
    Shell "cmd /k """ & strFile & """", strWndStyle
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    err.Number & vbCrLf & "Error Source: RunFile" & vbCrLf & "Error Description: " & _
    err.Description, vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Function combineTextFiles(ByRef returnFileList() As Variant, _
                          ByRef returnBankList() As Variant, _
                          ByVal fPath As String, _
                          ByVal aExt As String, _
                          ByVal newFileName As String) As String

'**** add windows scripting runtime library to your references

' Inputs: Path to folder to be scanned.
'         File extension of files to be combined.

If Left(aExt, 1) <> "." Then aExt = "." & aExt      ' Ensure that extension begins with a period.
If Right(fPath, 1) <> "\" Then fPath = fPath & "\"  ' Ensure that path name ends in a back slash.

Dim fName As Object, txtContent As String
Dim J As Long, I As Long, newFile As Object
Dim ii As Long: ii = -1 ' Set Initial value to properly increment subscript.
Dim jj As Long: jj = -1 ' Set Initial value to properly increment subscript.
Dim fs As Object
Dim farr() As Variant

'fArr = Array("download (8).QBO", "download (9).QBO", "download (10).QBO")
farr = getFileNames2(fPath, aExt)

Set fs = CreateObject("Scripting.filesystemobject")
'combineTextFiles = fPath & newFileName & "_" & Format(Now(), "yyyymmddhhmmss") & aExt
combineTextFiles = fPath & newFileName & "_" & Format(Now(), "yyyymmdd") & aExt
Set newFile = fs.createtextfile(combineTextFiles, True)

If Not IsArrayAllocated(farr) Then Exit Function
If UBound(farr) = 0 And farr(0) = "" Then Exit Function

Call QuickSort(farr, LBound(farr), UBound(farr))  ' Sort names of the files.

For J = LBound(farr) To UBound(farr)
    'Debug.Print ("fArr(" & J & ")=" & fArr(J))
    If Left(farr(J), Len(newFileName) + 3) = newFileName & "_20" Then GoTo Skip_File

    ii = ii + 1
    ReDim Preserve returnFileList(ii)
    returnFileList(ii) = fPath & farr(J)  ' Return back with name of found file.

    Set fName = fs.OpenTextFile(fPath & farr(J), ForReading)
        If fName.AtEndOfStream Then
            txtContent = ""
            Else
            txtContent = fName.ReadAll
        End If
        'fixFITID will ensure that FITID is not all numeric and will produce usable Unique ID.
        'Also extract and validate Acct and Bank routing numbers.  Alter acct number with AlphaID for security.
        
        Dim hAcctId As String
        Dim hBankId As String
        
        If aExt = ".qbo" Then _
          Call fixFITID(txtContent, _
                        hAcctId, _
                        hBankId, _
                        farr(J))
        newFile.Write txtContent & Chr(10) & Chr(13)
        ' Save list of Bank ID routing numbers encountered in the QBO files.
        
        Dim foundBankInList  As Boolean
        If IsArrayAllocated(returnBankList) Then
          For I = LBound(returnBankList) To UBound(returnBankList)
            If returnBankList(I) = hBankId Then foundBankInList = True
          Next I
        End If
        
        If Not foundBankInList Then
          jj = jj + 1
          ReDim Preserve returnBankList(jj)
          returnBankList(jj) = hBankId  ' Save unique Bank ID to return.
        End If

        
        
Skip_File:
Next J
Set fs = Nothing
End Function
Function getFileNames2(ByVal aPath As String, ByVal aExt As String) As Variant
'?getFileNames("G:\My Drive\Joel's Files\qbo files")
'Create an array of all files in folder with a given extension and return list of values.

If Left(aExt, 1) <> "." Then aExt = "." & aExt
If Right(aPath, 1) <> "\" Then aPath = aPath & "\"

Dim strFileName As String
'TODO: Specify path and file spec

'Dim strFolder As String: strFolder = "G:\My Drive\Joel's Files\qbo files\"
Dim strFolder As String: strFolder = aPath
If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"  ' Make sure that path ends in a \

Dim strFileSpec As String: strFileSpec = strFolder & "*" & aExt
Dim FileList() As Variant
Dim intFoundFiles As Long
strFileName = Dir(strFileSpec)
Do While Len(strFileName) > 0
    ReDim Preserve FileList(intFoundFiles)
    FileList(intFoundFiles) = strFileName
   'Debug.Print (FileList(intFoundFiles))
    intFoundFiles = intFoundFiles + 1
    strFileName = Dir
Loop
getFileNames2 = FileList()
End Function


Sub fixFITID(ByRef QBO_Str As String, _
             ByRef hAcctId As String, _
             ByRef hBankId As String, _
             ByVal FileName As String)

Debug.Print ("fixFITID is processing-" & FileName)
' Alter FITID element so it is not all numeric.
Dim I As Long, J As Long

'Get rid of leading zeros from BANK routing number.
Do While (InStr(1, QBO_Str, "<BANKID>0") <> 0)
  QBO_Str = replace(QBO_Str, "<BANKID>0", "<BANKID>")
Loop

'Get rid of leading zero from Bank Account Number.
Do While (InStr(1, QBO_Str, "<ACCTID>0") <> 0)
  QBO_Str = replace(QBO_Str, "<ACCTID>0", "<ACCTID>")
Loop

I = InStr(1, QBO_Str, "<FITID>20")
Do While I <> 0
  QBO_Str = Left(QBO_Str, I + 6) & "a" & Right(QBO_Str, Len(QBO_Str) - (I + 6))
  I = InStr(1, QBO_Str, "<FITID>20")
Loop _

Dim ErrorMsg   As String
'Make sure that <ACCTID> is found only one time.....
If strCount(QBO_Str, "<ACCTID>") > 1 Then
  ErrorMsg = "<ACCTID> expected in QBO file-" & FileName & " to be found 1 time only.  But was actually found " & _
        strCount(QBO_Str, "ACCTID") & " times." & _
        Chr(13) & Chr(13) & "Import process will be ended."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

'Make sure that <BANKID> is found only one time.....
If strCount(QBO_Str, "<BANKID>") > 1 Then
  ErrorMsg = "<BANKID> expected in QBO file-" & FileName & " to be found 1 time only.  But was actually found " & _
        strCount(QBO_Str, "BANKID") & " times." & _
        Chr(13) & Chr(13) & "Import process will be ended."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

' Now fix ACCTID to actually contain the account designator.
Dim acctId As String, bankId As String

hAcctId = ""
hBankId = ""
Dim hDigits   As String:  hDigits = "0123456789"
Dim pos  As Long

'Extract the <BANKID> into hBankdId and make sure it is numeric.......
I = InStr(1, QBO_Str, "<BANKID>")
If I = 0 Then
  hBankId = "9999999"   '  Default BANKID
  GoTo BankIDNotFound
End If
I = I + 8: pos = InStr(1, hDigits, Mid(QBO_Str, I, 1))
Do While (pos <> 0)
  hBankId = hBankId & Mid(QBO_Str, I, 1)
  I = I + 1: pos = InStr(1, hDigits, Mid(QBO_Str, I, 1))
Loop
If hBankId = "" Then
  ErrorMsg = "<BANKID> found in QBO file-" & FileName & " expected to be NUMERIC.  Found non-NUMERIC." & _
        Chr(13) & Chr(13) & "Import process will be ended."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If
BankIDNotFound:

Dim pAlphaID   As String
Dim pName_of_bank  As String
' Make sure this routing number is found in the Bank table and get the pAlphaID and pName_of_bank
Call EnsureValidBank(hBankId, FileName, pAlphaID, pName_of_bank)

'Extract the <ACCTID> into hAcctId and make sure it is numeric.......
J = InStr(1, QBO_Str, "<ACCTID>")
J = J + 8: pos = InStr(1, hDigits, Mid(QBO_Str, J, 1))
Do While (pos <> 0)
  hAcctId = hAcctId & Mid(QBO_Str, J, 1)
  J = J + 1: pos = InStr(1, hDigits, Mid(QBO_Str, J, 1))
Loop
If hAcctId = "" Then
  ErrorMsg = "<ACCTID> found in QBO file-" & FileName & " expected to be NUMERIC.  Found non-NUMERIC." & _
        Chr(13) & Chr(13) & "Import process will be ended."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

Dim HoldEndingBal   As Currency
Dim HoldEndingBalstr  As String:   HoldEndingBalstr = ""
I = InStr(1, QBO_Str, "<LEDGERBAL>")
J = InStr(I, QBO_Str, "<BALAMT>")
If I = 0 Or J = 0 Then
  ErrorMsg = "<LEDGERBAL><BALAMT> NOT found in QBO file-" & FileName & " as expected." & _
        Chr(13) & Chr(13) & "Import process will be ended."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If
For I = J + 8 To Len(QBO_Str)
  If Mid(QBO_Str, I, 1) = "<" Then Exit For
  HoldEndingBalstr = HoldEndingBalstr & Mid(QBO_Str, I, 1)
Next I
HoldEndingBal = HoldEndingBalstr


' Make sure this account number is linked to this routing number.
Call EnsureValidBankAccount(hBankId, hAcctId, FileName)

' For additional security, this will replace the AcctId with the alpha bank designator....
' Also add a file_name number to identify all records from a discrete file.
GBL_file_name = GBL_file_name + 1 ' Increment the file name by 1.
Dim hBankDesignator    As String
Dim hAcctIdStr         As String: hAcctIdStr = "<ACCTID>" & hAcctId
hBankDesignator = "<ACCTID>" & GBL_file_name & "~" & getBankDesignator(hBankId, hAcctId) & _
                  "~" & HoldEndingBal & "~" & FileName
If InStr(1, QBO_Str, "<BANKID>") = 0 Then _
  hBankDesignator = hBankDesignator & "<BANKID>" & hBankId
QBO_Str = replace(QBO_Str, hAcctIdStr, hBankDesignator)

End Sub
Sub EnsureValidBank(ByVal hBankId As String, _
                    ByVal FileName As String, _
                    ByRef pAlpha_id As String, _
                    ByRef pName_of_bank As String)

' Verify that rounting number can be found in the bank table.  If not, then stop the process.

Dim rst   As Recordset
Dim strSql    As String
Dim ErrorMsg  As String

strSql = "SELECT bank.name_of_bank, bank.routing_number, bank.alpha_id " _
       & "FROM bank WHERE (((bank.routing_number)=""" & hBankId & """));"
Debug.Print ("EnsureValidBank " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then GoTo Routing_Not_Found
pAlpha_id = rst!alpha_id
pName_of_bank = rst!name_of_bank
If Nz(rst!alpha_id, "") <> "" And Nz(rst!name_of_bank, "") <> "" Then
  rst.Close
  Exit Sub
End If

If Nz(rst!alpha_id, "") = "" Then
  ErrorMsg = "Bank Routing - " & hBankId & " has blank Alpha ID in the ""bank"" table." & Chr(13) & _
              "Process cannot continue until the Alpha ID is entered." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBank - Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

If Nz(rst!name_of_bank, "") = "" Then
  ErrorMsg = "Bank Routing - " & hBankId & " has blank Bank Name in the ""bank"" table." & Chr(13) & _
              "Process cannot continue until the Bank Name is entered." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBank - Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

Dim MsgResponse   As String

Routing_Not_Found:
ErrorMsg = "Bank Routing - " & hBankId & " was not found in the ""bank"" table." & Chr(13) & _
            "Process cannot continue until bank is entered into ""bank"" table." & Chr(13) & Chr(13) & _
            "You must remove file " & FileName & " from import folder or set up the new Bank Routing." & Chr(13) & Chr(13) & _
            "Do you want to program to set up a skeleton record? "
MsgResponse = MsgBox(ErrorMsg, vbYesNo, "EnsureValidBank - Routine....")
Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))

If MsgResponse = vbNo Then GoTo Abend

strSql = "INSERT INTO bank ( routing_number ) SELECT """ & hBankId & """ AS Expr1;"
DoCmd.RunSQL (strSql)
Debug.Print ("EnsureValidBank - " & strSql)

ErrorMsg = "Skeleton Bank Routing - " & hBankId & " has been set up." & Chr(13) & _
            "Now go and finish entering Bank Name and Bank ID, then restart process"
Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBank - Routine....")
Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
Abend:
  rst.Close
  End
  
  

End Sub

Sub EnsureValidBankAccount(ByVal hBankId As String, ByVal hAcctId As String, ByVal FileName As String)

' Make sure this account number is linked to this routing number.  If not, then stop the process or
' Set up Skeleton account before import process will continue.

Dim rst   As Recordset
Dim strSql    As String

Dim bank_id    As String
Dim bank_name  As String
'1)  Make sure that the routing number is good.
Call EnsureValidBank(hBankId, FileName, bank_id, bank_name)  ' Get the bank name.....

'2) Make sure that the combination of bank account number and routing number are found in the bank/bank_account table.
strSql = "SELECT bank_account.account_number, bank.alpha_id, bank.routing_number, bank.name_of_bank, bank_account.type_of_account, bank_account.bank_account_id, bank_account.Bank, bank_account.account_designator, bank_account.va_case_id, bank_account.depositor_account_title, bank_account.Fiduciary " _
       & "FROM bank_account INNER JOIN bank ON bank_account.Bank = bank.alpha_id " _
       & "WHERE (((bank_account.account_number)=""" & hAcctId & """) AND ((bank.routing_number)=""" & hBankId & """));"
Debug.Print ("EnsureValidBankAccount " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
  rst.Close
  GoTo Account_Not_Found
End If

'2a) Make sure that only one bank routing / bank account number combination is found.

'3)  Now ensure that all of the required fields are found in the bank account record.
If Nz(rst!bank_account_id, "") = "" Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has missing bank_account_id in the ""bank_account"" table." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3a) Check that Bank Alpha id is filled in.
If Nz(rst!alpha_id, "") = "" Then
  ErrorMsg = "Bank Routing - " & hBankId & " has blank Alpha ID in the ""bank"" table." & Chr(13) & _
              "Process cannot continue until the Alpha ID is entered." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3b) Check that the name of the bank is filled in.
If Nz(rst!name_of_bank, "") = "" Then
  ErrorMsg = "Bank Routing - " & hBankId & " has blank Bank Name in the ""bank"" table." & Chr(13) & _
              "Process cannot continue until the Bank Name is entered." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3c) Check that bank is filled in.
If Nz(rst!bank, "") = "" Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has blank Bank ID in the ""bank_account"" table." & Chr(13) & _
              "Process cannot continue until the ""bank_account"" is tied to a proper bank." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3d) Check that type_of_account is filled in.
If Nz(rst!type_of_account, "") <> "Checking" And _
   Nz(rst!type_of_account, "") <> "Savings" And _
   Nz(rst!type_of_account, "") <> "CreditCard" And _
   Nz(rst!type_of_account, "") <> "CD" Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has invalid type_of_account in the ""bank_account"" table." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3e) Check that account_designator is filled in.  (Not prefixed with "New Account")
If Nz(rst!account_designator, "") = "" Or _
   Left(Nz(rst!account_designator, ""), 11) = "New Account" Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has missing account_designator in the ""bank_account"" table." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3f) Check that va_case_id is filled in.
If Nz(rst!va_case_id, "") = "" Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has missing va_case_id in the ""bank_account"" table." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3g) Check that depositor_account_title is filled in.
If Nz(rst!depositor_account_title, "") = "" Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has missing depositor_account_title in the ""bank_account"" table." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

'3g) Check that Fiduciary is filled in.
If Nz(rst!Fiduciary, 0) = 0 Then
  ErrorMsg = hAcctId & "/" & hBankId & " - " & rst!account_designator & " has missing fiduciary in the ""bank_account"" table." & Chr(13) & Chr(13) & _
              "Routine will abort. Restart import process after bank/bank account tables are corrected." & Chr(13) & Chr(13)
  Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  GoTo Abend
End If

GoTo Good_return

Dim MsgResponse   As String

'Account_Not_Found:

'3c) Check to see if the account is good, but not matched the routing number found in the bank ID.


strSql = "SELECT bank_account.account_number, bank.routing_number, bank.name_of_bank, bank_account.bank_account_id, bank_account.Bank, bank_account.account_designator, bank_account.va_case_id, bank_account.depositor_account_title, bank_account.Fiduciary " _
       & "FROM bank_account LEFT JOIN bank ON bank_account.Bank = bank.alpha_id " _
       & "WHERE (((bank_account.account_number)=" & hAcctId & "));"
Debug.Print ("EnsureValidBankAccount " & strSql)
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then




  rst.Close
  GoTo Account_Not_Found
End If

Account_Not_Found:

ErrorMsg = "Account-" & hAcctId & " for Bank - " & hBankId & "/" & bank_name & " was not found in the ""bank_account"" table." & Chr(13) & Chr(13) & _
            "Process cannot continue until account is entered into ""bank_account"" table." & Chr(13) & Chr(13) & _
            "You must either remove file " & FileName & " from import folder or set up the new Bank Account." & Chr(13) & Chr(13) & _
            "Do you want to program to set up a skeleton account record? "
MsgResponse = MsgBox(ErrorMsg, vbYesNo, "EnsureValidBankAccount- Routine....")
Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))

If MsgResponse = vbNo Then GoTo Abend

strSql = "INSERT INTO bank_account ( account_number, Bank, account_designator ) " _
       & "SELECT " & hAcctId & " AS Expr1, """ & bank_id & """ AS Expr2, ""REPLACE New Account for " & bank_name & """ AS Expr3;"
Debug.Print ("EnsureValidBankAccount " & strSql)
DoCmd.RunSQL (strSql)

ErrorMsg = "Skeleton Bank Account-" & hAcctId & " for Bank - " & hBankId & "/" & bank_name & " has been set up." & Chr(13) & Chr(13) & _
            "Now go and finish entering Bank Account ID and other missing fields, then restart process"
Call MsgBox(ErrorMsg, vbOKOnly, "EnsureValidBankAccount- Routine....")
Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
Abend:
  End
  
Good_return:
  rst.Close
  Exit Sub

End Sub

Function replacex()
Dim str1  As String
str1 = "One fish, two fish, red fish, blue fish"
Debug.Print (str1)
str1 = replace(str1, "fish", "cat")

Debug.Print (str1)
End Function


Function getBankDesignator(ByVal routingNumber As String, ByVal acctNumber As String) As String

' Find account designator to be used when replacing the account number in the QBO file.

'acctNumber = "97599526"
'routingNumber = "62204530"

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

strSql = "SELECT bank.routing_number, bank_account.account_number, bank_account.account_designator " _
       & "FROM bank INNER JOIN bank_account ON bank.alpha_id = bank_account.Bank " _
       & "WHERE (((bank.routing_number)= """ & routingNumber & """) AND ((bank_account.account_number)=""" & acctNumber & """)); "
Debug.Print ("getBankDesignator " & strSql)

getBankDesignator = "Unknown"
For I = Len(acctNumber) To 1 Step -1
  getBankDesignator = getBankDesignator & Mid(acctNumber, I, 1)
Next I

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
  rst.Close
  Exit Function
End If

If rst.RecordCount > 1 Then
  ErrorMsg = "Account- " & acctNumber & "Duplicate bank accounts found in the bank_account table.  This is not allowed.  Must be resolved before proceeding." & _
        Chr(13) & Chr(13) & "Import process will be ended."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

getBankDesignator = rst!account_designator

rst.Close

End Function

Function strCount(ByVal xx As String, ByVal findStr As String) As Long

Dim I         As Long
Dim xstart    As Long: xstart = 1

I = InStr(xstart, xx, findStr)
Do While (I <> 0)
  strCount = strCount + 1
  xstart = I + 1
  I = InStr(xstart, xx, findStr)
Loop
End Function

Function getFileNames(ByVal pFileName As String) As Variant
 
'Create an array of all files in folder within a given pattern............
 
Dim I As Long
I = InStrRev(pFileName, "\")
If I = 0 Then
  ErrorMsg = "Invalid Path name """ & pFileName & """  passed to getFileNames." & _
              Chr(13) & Chr(13) & "Process will abort..........."
  Call MsgBox(ErrorMsg, vbOKOnly, "getFileNames *******************")
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

Dim aFolder As String
aFolder = Left(pFileName, I)

Dim FileList() As Variant
Dim intFoundFiles As Long
pFileName = Dir(pFileName)

Do While Len(pFileName) > 0
    ReDim Preserve FileList(intFoundFiles)
    FileList(intFoundFiles) = aFolder & pFileName
    intFoundFiles = intFoundFiles + 1
    pFileName = Dir
Loop
getFileNames = FileList()
End Function

Public Function RenameFileOrDir(ByVal strSource As String, ByVal strTarget As String, _
  Optional fOverwriteTarget As Boolean = False) As Boolean
 
  On Error GoTo PROC_ERR
 
  Dim fRenameOK As Boolean
  Dim fRemoveTarget As Boolean
  Dim strFirstDrive As String
  Dim strSecondDrive As String
  Dim fOK As Boolean
 
  If Not ((Len(strSource) = 0) Or (Len(strTarget) = 0) Or (Not (FileOrDirExists(strSource)))) Then
 
    ' Check if the target exists
    If FileOrDirExists(strTarget) Then
 
      If fOverwriteTarget Then
        fRemoveTarget = True
      Else
        If MsgBox("Do you wish to overwrite the target file?", vbExclamation + vbYesNo, "Overwrite confirmation") = vbYes Then
          fRemoveTarget = True
        End If
      End If
 
      If fRemoveTarget Then
        ' Check that it's not a directory
        If ((GetAttr(strTarget) And vbDirectory)) <> vbDirectory Then
          Kill strTarget
          fRenameOK = True
        Else
          MsgBox "Cannot overwrite a directory", vbOKOnly, "Cannot perform operation"
        End If
      End If
    Else
      ' The target does not exist
      ' Check if source is a directory
      If ((GetAttr(strSource) And vbDirectory) = vbDirectory) Then
        ' Source is a directory, see if drives are the same
        strFirstDrive = Left(strSource, InStr(strSource, ":\"))
        strSecondDrive = Left(strTarget, InStr(strTarget, ":\"))
        If strFirstDrive = strSecondDrive Then
          fRenameOK = True
        Else
          MsgBox "Cannot rename directories across drives", vbOKOnly, "Cannot perform operation"
        End If
      Else
        ' It's a file, ok to proceed
        fRenameOK = True
      End If
    End If
 
    If fRenameOK Then
      Name strSource As strTarget
      fOK = True
    End If
  End If
 
  RenameFileOrDir = fOK
 
PROC_EXIT:
  Exit Function
 
PROC_ERR:
  MsgBox "Error: " & err.Number & ". " & err.Description, , "RenameFileOrDir"
  Resume PROC_EXIT
End Function

Public Function FileOrDirExists(strDest As String) As Boolean
  Dim intLen As Long
  Dim fReturn As Boolean

  fReturn = False

  If strDest <> vbNullString Then
    On Error Resume Next
    intLen = Len(Dir$(strDest, vbDirectory + vbNormal))
    On Error GoTo PROC_ERR
    fReturn = (Not err And intLen > 0)
  End If

PROC_EXIT:
  FileOrDirExists = fReturn
  Exit Function

PROC_ERR:
  MsgBox "Error: " & err.Number & ". " & err.Description, , "FileOrDirExists"
  Resume PROC_EXIT
End Function

Public Sub MakeMyFolder(FolderAndPath As String)
'Updateby Extendoffice 20161109
    Dim fdObj As Object

   ' Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    If Not fdObj.FolderExists(FolderAndPath) Then
        fdObj.CreateFolder (FolderAndPath)
        MsgBox "Git folder-" & FolderAndPath & " has been created.", vbInformation
    End If
   ' Application.ScreenUpdating = True
End Sub

