Attribute VB_Name = "SharedQBO"
Option Compare Database
Option Explicit

Dim ErrorMsg   As String

Public Sub buildQboTable(pBank As String)

Dim tdf       As TableDef
Dim target    As TableDef
Dim db As dao.Database, rst As Recordset
Dim fld As dao.Field
Dim prop As dao.Property
Dim I As Long, J As Long, myDataType As Long

Call Create_A_Table(pBank)

Set db = CurrentDb
Set tdf = db.TableDefs(pBank)

'  Text  10
'  dbDate  8
'  dbDouble 7

'  Fields added to the QBO file
tdf.Fields.Append tdf.CreateField("account_designator", 10)   ' Text
tdf.Fields.Append tdf.CreateField("file_number", 7)   ' dbDouble
tdf.Fields.Append tdf.CreateField("qbo_file_name", 10)   ' Text
tdf.Fields.Append tdf.CreateField("ending_balance", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("va_case_id", 10)   ' Text
tdf.Fields.Append tdf.CreateField("missing_trans", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("prev_date_posted", 8)   ' Date

' Fields from the QBO file
tdf.Fields.Append tdf.CreateField("balance", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("save_balance", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("amount", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("row_number", 7)
tdf.Fields.Append tdf.CreateField("date_posted", 8)   ' Date
tdf.Fields.Append tdf.CreateField("payee", 10)   ' Text
tdf.Fields.Append tdf.CreateField("in", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("out", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("currency", 10)   ' Text
tdf.Fields.Append tdf.CreateField("memo", 10)   ' Text
tdf.Fields.Append tdf.CreateField("check_number", 10)   ' dbDouble
tdf.Fields.Append tdf.CreateField("unique_transaction_id", 10)   ' Text
tdf.Fields.Append tdf.CreateField("record_type", 10)   ' Text
tdf.Fields.Append tdf.CreateField("investment_action", 10)   ' Text
tdf.Fields.Append tdf.CreateField("security_id", 10)   ' Text
tdf.Fields.Append tdf.CreateField("security_name", 10)   ' Text
tdf.Fields.Append tdf.CreateField("ticker", 10)   ' Text
tdf.Fields.Append tdf.CreateField("price", 5)   ' dbCurrency
tdf.Fields.Append tdf.CreateField("quantity_of_shares", 7)   ' dbDouble
tdf.Fields.Append tdf.CreateField("commission", 7)   ' dbDouble
tdf.Fields.Append tdf.CreateField("trade_date", 8)   ' Date
tdf.Fields.Append tdf.CreateField("sell_type", 10)   ' Text
tdf.Fields.Append tdf.CreateField("buy_type", 10)   ' Text
tdf.Fields.Append tdf.CreateField("initiated", 8)   ' Date
tdf.Fields.Append tdf.CreateField("settle_date", 8)   ' Date
tdf.Fields.Append tdf.CreateField("account_number", 10)   ' dbDouble
tdf.Fields.Append tdf.CreateField("account_type", 10)   ' Text
tdf.Fields.Append tdf.CreateField("bank_id", 10)   ' dbDouble
tdf.Fields.Append tdf.CreateField("branch_id", 7)   ' dbDouble
tdf.Fields.Append tdf.CreateField("fi_org", 10)   ' Text
tdf.Fields.Append tdf.CreateField("fi_id", 10)   ' Text
tdf.Fields.Append tdf.CreateField("intu_bid", 10)   ' Text
tdf.Fields.Append tdf.CreateField("filename", 10)   ' Text

Set fld = Nothing
Set tdf = Nothing
Set target = Nothing
Set db = Nothing
Set prop = Nothing

End Sub

Public Sub exportToExcel(ByVal fPath As String, _
                          ByVal newFileName As String, _
                          ByVal ext As String, _
                          ByVal BankDbPrefix As String, _
                          ByRef outputfilename As String)

If Right(fPath, 1) <> "\" Then fPath = fPath & "\"  ' Ensure that path name ends in a back slash.
If Left(ext, 1) <> "." Then ext = "." & ext

'Before exporting to excel, call Remove_Table_Field to remove the "excel_row_number" field....
Call Remove_Table_Field("excel_row_number", BankDbPrefix)

outputfilename = fPath & newFileName & "_NEWQBO_" & Format(Now(), "yyyymmddhhmmss") & ext
'Debug.Print ("outputFileName=" & outputFileName)
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, BankDbPrefix, outputfilename, True, "QBO"

End Sub

Public Function integrityCheckQBO(ByVal BankDbPrefix As String, Optional ByVal pHeader As String) As String

' This routine will validate various parts of the QBO file received from the bank.

If pHeader = "" Then pHeader = "General QBO Integrity Check "
If Right(pHeader, 1) <> " " Then pHeader = pHeader & " "
pHeader = pHeader & "***************************************"

'1) Bank account is found in the "bank_account" database table.
    Call confirmValidAccount(BankDbPrefix, pHeader)
    
'1a) Bank account Designator is found to be a valid value not starting with REPLACE.
    Call confirmValidAccountDesignator(BankDbPrefix, pHeader)
    
'1b) Verify that there are no gaps in row_number values for a file.  Ensure no missing records.
    Call confirmNoMissingRows(BankDbPrefix, pHeader)
    
'2) Routing number matches bank routing number found the "bank" database table.
    Call confirmBankRouting(BankDbPrefix, pHeader)

'3) Running balance has integrity and is in balance with itself.
    Call ensureRunningBalanceIntegrity(BankDbPrefix, pHeader)

'4) All "account_designator" fields are filled in and match with the "bank_account" database table.
    Call ensureBankAccountsProperlyLinked(BankDbPrefix, pHeader)
    
'4a) All fields have a valid veteran case number.
    Call ensure_ValidCaseNumber(BankDbPrefix, pHeader)
    
'4b) Verify that there are no date/balance/amount mis-matches with balances in "transaction" table.
    Call confirmNoMissMatches(BankDbPrefix, pHeader)
    
    

'5) All "account_designator" fields found in the "transaction" table have integrity with the "bank_account", before adding
'   new records, make sure that the bank account tables have not been changed to mis-match what is in the QBO table.

  Exit Function
End Function


Private Sub confirmValidAccount(ByVal BankDbPrefix As String, ByVal pHeader As String)

'1) Bank account is found in the "bank_account" database table.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

strSql = "SELECT " & BankDbPrefix & ".account_number AS qbo_acct, bank_account.account_number AS bank_account_number, bank.routing_number, bank.name_of_bank " _
       & "FROM (" & BankDbPrefix & " LEFT JOIN bank_account ON " & BankDbPrefix & ".account_number = bank_account.account_number) INNER JOIN bank ON " & BankDbPrefix & ".bank_id = bank.routing_number " _
       & "GROUP BY " & BankDbPrefix & ".account_number, bank_account.account_number, bank.routing_number, bank.name_of_bank;"
       
Debug.Print ("Integrity Check 1) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all routing numbers are found in the "bank" table.
Do
  If Nz(rst!qbo_acct, 0) <> Nz(rst!bank_account_number, 0) Then GoTo Report_Error
  rst.MoveNext
  If rst.EOF Then Exit Sub
Loop

Report_Error:
rst.MoveFirst
Dim displayMsg   As String
displayMsg = "Unknown Bank Account numbers found in QBO file." & Chr(13) & Chr(13) & _
             "Not found in ""bank_account"" table:" & Chr(13)
Do
  If Nz(rst!qbo_acct, 0) <> Nz(rst!bank_account_number, 0) Then _
    displayMsg = displayMsg & "Bank Account - " & rst!qbo_acct & "    " & _
                 Nz(rst!routing_number, "") & " " & Nz(rst!name_of_bank, "UnKnown Bank") & Chr(13)
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
  displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
  Call MsgBox(displayMsg, vbOKOnly, pHeader)
  Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
  End

rst.Close

End Sub


Private Sub confirmValidAccountDesignator(ByVal BankDbPrefix As String, ByVal pHeader As String)

'1a) Bank account Designator is found to be a valid value not starting with REPLACE.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

strSql = "SELECT " & BankDbPrefix & ".account_number AS qbo_acct, " _
              & "" & BankDbPrefix & ".account_designator AS qbo_acct_designator, " _
              & "bank_account.account_number AS bank_account_number, " _
              & "bank.routing_number, bank.name_of_bank " _
       & "FROM (" & BankDbPrefix & " LEFT JOIN bank_account ON " & BankDbPrefix & ".account_number = bank_account.account_number) INNER JOIN bank ON " & BankDbPrefix & ".bank_id = bank.routing_number " _
       & "GROUP BY " & BankDbPrefix & ".account_number, " _
                & "" & BankDbPrefix & ".account_designator, " _
                & "bank_account.account_number, " _
                & "bank.routing_number, bank.name_of_bank;"

Debug.Print ("Integrity Check 1a) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all routing numbers are found in the "bank" table.
Do
  If Nz(rst!qbo_acct_designator, "") = "" Then GoTo Report_Error
  If Left(Nz(rst!qbo_acct_designator, ""), 7) = "REPLACE" Then GoTo Report_Error
  rst.MoveNext
  If rst.EOF Then Exit Sub
Loop

Report_Error:
rst.MoveFirst
Dim displayMsg   As String
displayMsg = "New skeleton Account Designators found have not been filled in the Bank Account table for:" & Chr(13) & Chr(13)
Do
  If (Nz(rst!qbo_acct_designator, "") = "") Or (Left(Nz(rst!qbo_acct_designator, ""), 7) = "REPLACE") Then _
    displayMsg = displayMsg & "Bank Account-" & rst!qbo_acct & "/" & _
                 Nz(rst!routing_number, "") & "  Designator-" & Nz(rst!qbo_acct_designator, "UnKnown") & Chr(13) & Chr(13)
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
  displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
  Call MsgBox(displayMsg, vbOKOnly, pHeader)
  Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
  End

rst.Close

End Sub


Private Sub confirmNoMissingRows(ByVal BankDbPrefix As String, ByVal pHeader As String)

'1b) Verify that there are no gaps in row_number values for a file.  Ensure no missing records.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset
Dim LastRowNumber  As Double:   LastRowNumber = 0
Dim LastFileNumber As Double:   LastFileNumber = 0
Dim FoundError   As Boolean:    FoundError = False

strSql = "SELECT " & BankDbPrefix & ".* " _
       & "FROM " & BankDbPrefix & " " _
       & "ORDER BY " & BankDbPrefix & ".row_number;"

Debug.Print ("Integrity Check 1b) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub

'  Now loop thru the query results and report on the mis-matches.

LastRowNumber = 0: LastFileNumber = 0 ' Initial values.....

rst.MoveFirst
Dim displayMsg   As String
displayMsg = "Missing Row number(s) found in union_state_temp_qbo table: " & Chr(13) & Chr(13)
Do
  If rst!file_number <> LastFileNumber Then
    LastFileNumber = rst!file_number
    LastRowNumber = rst!row_number - 1
  End If
  
  If LastRowNumber + 1 <> rst!row_number Then
    FoundError = True
    displayMsg = displayMsg & _
       (rst!row_number - (LastRowNumber + 1)) & " records missing after row_number-" & LastRowNumber & _
       " for " & rst!account_designator & Chr(13)
    LastRowNumber = rst!row_number - 1
  End If
  LastRowNumber = LastRowNumber + 1
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
  rst.Close
  If Not FoundError Then Exit Sub
  displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
  Call MsgBox(displayMsg, vbOKOnly, pHeader)
  Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
  End



End Sub



Private Sub confirmBankRouting(ByVal BankDbPrefix As String, ByVal pHeader As String)

'2) Routing number matches bank routing number found the "bank" database table.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset
Dim ErrorMsg   As String

strSql = "SELECT " & BankDbPrefix & ".bank_id, bank.routing_number, bank.name_of_bank " _
       & "FROM " & BankDbPrefix & " LEFT JOIN bank ON " & BankDbPrefix & ".bank_id = bank.routing_number " _
       & "GROUP BY " & BankDbPrefix & ".bank_id, bank.routing_number, bank.name_of_bank;"
Debug.Print ("Integrity Check 2) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all routing numbers are found in the "bank" table.
Do
  If Nz(rst!routing_number, 0) = 0 Then
    ErrorMsg = "Routing Number (" & rst!bank_id & ") was not found in the ""bank"" table." & Chr(13) _
        & Chr(13) & "Import Process will be ABORTED....."
    Call MsgBox(ErrorMsg, vbOKOnly, pHeader)
    Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
    End
  End If
  rst.MoveNext
  If rst.EOF Then
    rst.Close
    Exit Sub
  End If
Loop



End Sub

Sub ensureBankAccountsProperlyLinked(ByVal BankDbPrefix As String, ByVal pHeader As String)

'4) All "account_designator" fields are filled in and match with the "bank_account" database table.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

strSql = "SELECT " & BankDbPrefix & ".account_designator, " & BankDbPrefix & ".va_case_id, " & BankDbPrefix & ".account_number, " & BankDbPrefix & ".account_type, bank.alpha_id, " & BankDbPrefix & ".bank_id, bank.name_of_bank, bank_account.bank_account_id, bank_account.account_designator " _
       & "FROM (" & BankDbPrefix & " INNER JOIN bank ON " & BankDbPrefix & ".bank_id = bank.routing_number) INNER JOIN bank_account ON " & BankDbPrefix & ".account_number = bank_account.account_number " _
       & "GROUP BY " & BankDbPrefix & ".account_designator, " & BankDbPrefix & ".va_case_id, " & BankDbPrefix & ".account_number, " & BankDbPrefix & ".account_type, bank.alpha_id, " & BankDbPrefix & ".bank_id, bank.name_of_bank, bank_account.bank_account_id, bank_account.account_designator " _
       & "HAVING (((" & BankDbPrefix & ".account_designator) Is Null)) OR (((" & BankDbPrefix & ".account_designator)=''));"

Debug.Print ("Integrity Check 4) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all account_designator(s) are filled in.....
Dim displayMsg   As String
displayMsg = "account_designator(s) found in QBO file are BLANK." & Chr(13) & Chr(13) & _
             "Errors found in ""bank_account"" table:" & Chr(13) & Chr(13)
Do
  displayMsg = displayMsg & "Bank Account - " & Nz(rst!bank_account_id, "") & "-" & Nz(rst!account_number, "") & "  and  " & Nz(rst!alpha_id, "") & _
                 "-" & Nz(rst!bank_id, "") & " " & Nz(rst!name_of_bank, "UnKnown Bank") & " are not properly linked." & Chr(13)
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
  displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
  Call MsgBox(displayMsg, vbOKOnly, pHeader)
  Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
  End

rst.Close

End Sub

Sub ensure_ValidCaseNumber(ByVal BankDbPrefix As String, ByVal pHeader As String)

'4a) All fields have a valid veteran case number.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

strSql = "SELECT " & BankDbPrefix & ".va_case_id as qbo_va_case_id, " & BankDbPrefix & ".account_number, " & BankDbPrefix & ".account_type, " & BankDbPrefix & ".bank_id, veteran.va_case_id, bank.alpha_id, bank.name_of_bank, bank_account.bank_account_id " _
       & "FROM ((" & BankDbPrefix & " LEFT JOIN veteran ON " & BankDbPrefix & ".va_case_id = veteran.va_case_id) INNER JOIN bank ON " & BankDbPrefix & ".bank_id = bank.routing_number) INNER JOIN bank_account ON " & BankDbPrefix & ".account_number = bank_account.account_number " _
       & "GROUP BY " & BankDbPrefix & ".va_case_id, " & BankDbPrefix & ".account_number, " & BankDbPrefix & ".account_type, " & BankDbPrefix & ".bank_id, veteran.va_case_id, bank.alpha_id, bank.name_of_bank, bank_account.bank_account_id " _
       & "HAVING (((veteran.va_case_id) Is Null Or (veteran.va_case_id)=''));"
Debug.Print ("Integrity Check 4a) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all account_designator(s) are filled in.....
Dim displayMsg   As String
displayMsg = "Veteran Case Number(s) found in QBO file are not found in ""veteran"" table." & _
             "  Likely cause is incorrect veteran linked to bank account." & Chr(13) & Chr(13) & _
             "Errors found:" & Chr(13) & Chr(13)
Do
  displayMsg = displayMsg & "Bank Account - " & Nz(rst!bank_account_id, "") & "-" & Nz(rst!account_number, "") & " is improperly linked to Veteran-" & _
                 Nz(rst!qbo_va_case_id, "") & Chr(13) & Chr(13)
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
  displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
  Call MsgBox(displayMsg, vbOKOnly, pHeader)
  Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
  End

rst.Close

End Sub


Private Sub ensureRunningBalanceIntegrity(ByVal BankDbPrefix As String, ByVal pHeader As String)

'3) Running balance has integrity and is in balance with itself.

Dim strSql              As String
Dim rst                 As Recordset
Dim runningBalance      As Double
Dim missingTransactions As Double
Dim thisKey             As String, lastKey As String: lastKey = "???"
Dim holdPrevDate        As Date

strSql = "UPDATE " & BankDbPrefix & " SET " & BankDbPrefix & ".missing_trans = 0;"
DoCmd.RunSQL (strSql)
strSql = "UPDATE " & BankDbPrefix & " SET " & BankDbPrefix & ".prev_date_posted = NULL;"
DoCmd.RunSQL (strSql)

strSql = "SELECT " & BankDbPrefix & ".account_designator, " & BankDbPrefix & ".file_number, " & BankDbPrefix & ".balance, " & BankDbPrefix & ".amount, " & BankDbPrefix & ".row_number, " & BankDbPrefix & ".missing_trans, " & BankDbPrefix & ".prev_date_posted, " & BankDbPrefix & ".date_posted " _
       & "FROM " & BankDbPrefix & " " _
       & "ORDER BY " & BankDbPrefix & ".file_number, " & BankDbPrefix & ".date_posted, " & BankDbPrefix & ".amount DESC , " & BankDbPrefix & ".unique_transaction_id DESC;"
       
Debug.Print ("Integrity Check 3-a) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql) ' Open recordset with intent to edit.
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all account_designator(s) are filled in.....
Dim displayMsg   As String
Do
  thisKey = rst!file_number
  If thisKey <> lastKey Then runningBalance = rst!balance - rst!amount
  lastKey = thisKey
  runningBalance = runningBalance + rst!amount
  
  missingTransactions = runningBalance - rst!balance
  If missingTransactions <> 0 Then
    rst.Edit
    rst!missing_trans = missingTransactions * -1
    rst!prev_date_posted = holdPrevDate
    runningBalance = runningBalance + rst!missing_trans
    rst.Update
  End If
  
  holdPrevDate = rst!date_posted
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
rst.Close

' Now check to see if any transactions are missing.....
strSql = "SELECT " & BankDbPrefix & ".account_designator, " & BankDbPrefix & ".file_number, " & BankDbPrefix & ".prev_date_posted, " & BankDbPrefix & ".missing_trans, " & BankDbPrefix & ".date_posted, " & BankDbPrefix & ".amount, " & BankDbPrefix & ".payee " _
       & "FROM " & BankDbPrefix & " " _
       & "WHERE (((" & BankDbPrefix & ".missing_trans)<>0));"
       
Debug.Print ("Integrity Check 3-b) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all account_designator(s) are filled in.....

displayMsg = "Integrity Check found transaction(s) missing from QBO import. " & Chr(13) & Chr(13) & _
             "Missing Errors found:" & Chr(13) & Chr(13)

Do
  displayMsg = displayMsg & "Account - " & Nz(rst!account_designator, "") & " " & Nz(rst!missing_trans, 0) & " is missing between " & _
               Nz(rst!prev_date_posted, "") & " and " & Nz(rst!date_posted, "") & Chr(13) & Chr(13)
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop3
Loop
Finished_Do_Loop3:
  displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
  Call MsgBox(displayMsg, vbOKOnly, pHeader)
  Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
  End

rst.Close

End Sub


Public Function SelectFileQBO(ByRef HoldFileName As String, ByRef HoldFilePath As String, ByRef HoldSelItem As String, Optional ByVal preSelectedFile As String = "")
   'Dim fd As Office.FileDialog
   'Set fd = Application.FileDialog(msoFileDialogFilePicker)
   Dim fd As Object
   Dim pos   As Long
   Dim Response  '  Yes=6  No=7  Retry=4  OK=1  Cancel=2
   Dim askUserForFileSelection    As Boolean
   askUserForFileSelection = True
   If preSelectedFile <> "" Then askUserForFileSelection = False
   
Retry_Selection:
   HoldSelItem = ""   ' Clear last selection....
   If Not askUserForFileSelection Then
     HoldSelItem = preSelectedFile
     askUserForFileSelection = True
     GoTo File_is_Selected
   End If
   '  Set Initial returning values
   
   Set fd = Application.FileDialog(1)
   If Nz(HoldFilePath, "") = "" Then HoldFilePath = CurrentProject.Path
   With fd
      .InitialFileName = "" & HoldFilePath & "\*.xlsx"
      .Title = "Select a File"
      .Filters.Clear
      .Filters.Add "Excel Files", "*.xlsx"
      If .Show Then HoldSelItem = .SelectedItems(1)
File_is_Selected:
      
      SelectFileQBO = HoldSelItem   '  This is the whole path and file name.
      
      If HoldSelItem = "" Then Exit Function
     '  Parse out the name of the file....
      pos = InStrRev(HoldSelItem, "\")
      If pos <> 0 And pos <> Len(HoldSelItem) And pos <> 1 Then
         SelectFileQBO = Mid(HoldSelItem, pos + 1)
         HoldFileName = Mid(HoldSelItem, pos + 1)
         HoldFilePath = Mid(HoldSelItem, 1, pos - 1)
      End If
      
   End With
   
   If Right(SelectFileQBO, 4) <> ".qbo" Then
      
      ErrorMsg = "INVALID file-'" & SelectFileQBO & "' was chosen." & Chr(13) & Chr(13) _
        & "Must be *.qbo to be valid." & Chr(13) & Chr(13) _
        & "Retry to select another file,  Cancel to Quit."
      Response = MsgBox(ErrorMsg, vbRetryCancel)
      Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
      If Response = 4 Then  ' Retry Button
        GoTo Retry_Selection
        Else
          SelectFileQBO = ""
          HoldFileName = ""
          Set fd = Nothing
          Exit Function
      End If
      Exit Function
   End If
   

   Set fd = Nothing
End Function

Public Sub CreateAndRunProperSoftBATfile(ByRef HoldFilePath As String, _
                                  ByVal aInputFile As String, _
                                  ByVal aPgm As String, _
                                  ByVal aBank As String, _
                                  ByVal ext As String)

Dim aLine As String
Dim holdCmd   As String

Dim fs As Object, a As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.createtextfile(HoldFilePath & "\" & aBank & aPgm & ".bat", True)
a.WriteLine ("PATH C:\Program Files (x86)\ProperSoft\" & aPgm)
aLine = aPgm & ".exe """ & aInputFile & """ """ & HoldFilePath & "\" & aBank & ext & ".csv"""
Debug.Print ("aLine=" & aLine)

a.WriteLine (aLine)
a.WriteLine ("exit")
a.Close
Set fs = Nothing
Set a = Nothing
    
holdCmd = HoldFilePath & "\" & aBank & aPgm & ".bat"
Debug.Print (holdCmd)
Call RunFile(holdCmd, vbHide)

Dim waitTill As Date
waitTill = Now() + TimeValue("00:00:10")  ' Wait 10 seconds for program to finish....
While Now() < waitTill
  DoEvents
Wend

End Sub


Public Function cleanUpQboFiles(ByVal HoldFilePath As String, _
                                ByVal transactionsBackupName As String, _
                                ByVal BankDbPrefix As String, _
                                ByVal BankFilePrefix As String)
    Dim I As Long
    Dim HoldTransBkupName As String
    
    I = InStrRev(transactionsBackupName, "_")
    HoldTransBkupName = Left(transactionsBackupName, I)

    If HoldTransBkupName <> "" Then Call DelTblS(HoldTransBkupName & "*") ' Delete all backup tables.
    If BankDbPrefix <> "" Then Call DelTbl(BankDbPrefix)
    Dim fArrKillFileList() As Variant
    fArrKillFileList = getFileNames(HoldFilePath & "\" & BankFilePrefix & "*.*")
    
    'combinedPathAndFile
    If fArrKillFileList(0) <> "" Then
      For I = LBound(fArrKillFileList) To UBound(fArrKillFileList)
        Call Kill(fArrKillFileList(I))
      Next I
    End If
   
End Function

Private Sub confirmNoMissMatches(ByVal BankDbPrefix As String, ByVal pHeader As String)

'4b) Verify that there are no date/balance/amount mis-matches with balances in "transaction" table.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset
Dim LastRowNumber  As Double:   LastRowNumber = 0
Dim LastFileNumber As Double:   LastFileNumber = 0
Dim FoundError   As Boolean:    FoundError = False

strSql = "SELECT union_state_temp_qbo.id, union_state_temp_qbo.account_designator, union_state_temp_qbo.file_number, union_state_temp_qbo.qbo_file_name, union_state_temp_qbo.date_posted, transactions.posted_date, union_state_temp_qbo.amount as qbo_amt, transactions.Amount as trans_amt, union_state_temp_qbo.balance, transactions.running_balance, transactions.account_number, transactions.qbo_unique_trans_id " _
       & "FROM union_state_temp_qbo INNER JOIN transactions ON (union_state_temp_qbo.unique_transaction_id = transactions.qbo_unique_trans_id) AND (union_state_temp_qbo.account_number = transactions.account_number) " _
       & "WHERE (((union_state_temp_qbo.date_posted<>[transactions].[posted_date]) or " _
       & "(union_state_temp_qbo.amount<>[transactions].[Amount]) or " _
       & "(union_state_temp_qbo.balance<>[transactions].[running_balance])) AND ((transactions.qbo_unique_trans_id) Is Not Null Or (transactions.qbo_unique_trans_id)<>""""));"

Debug.Print ("Integrity Check 4b) " & strSql)


Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub  ' No exceptions were found....

'  Now loop thru the query results and report on the mis-matches.

Dim displayMsg   As String
displayMsg = "Integrity check compare of QBO and transactions table found following MisMatches on overlapping records: " & Chr(13) & Chr(13)
Do
  displayMsg = displayMsg & "File-" & rst!qbo_file_name & " QBO Id-" & rst!Id & Chr(13)
  If rst!date_posted <> rst!posted_date Then _
    displayMsg = displayMsg & "        QBO=" & rst!date_posted & "  Transactions=" & rst!posted_date & Chr(13)
  If rst!qbo_amt <> rst!trans_amt Then _
    displayMsg = displayMsg & "        QBO=" & Format(rst!qbo_amt, "currency") & "  Transactions=" & Format(rst!trans_amt, "Currency") & Chr(13)
  If rst!balance <> rst!running_balance Then _
    displayMsg = displayMsg & "        QBO-Bal=" & Format(rst!balance, "currency") & "  Transactions-Bal=" & Format(rst!running_balance, "Currency") & Chr(13)
  rst.MoveNext
  If rst.EOF Then Exit Do
Loop
  
rst.Close
displayMsg = displayMsg & Chr(13) & "Import Process will be ABORTED....."
Call MsgBox(displayMsg, vbOKOnly, pHeader)
Debug.Print (Chr(13) & "****" & displayMsg & Chr(13))
End

End Sub


Public Sub ParseOutFileNumber(ByVal BankDbPrefix As String)

'4c  Parse out the file_number from the account_designator field.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

Dim FileNumber  As String
Dim AccountDesignator As String
Dim EndingBalance  As String
Dim FileName As String

strSql = "SELECT " & BankDbPrefix & ".account_designator, " & BankDbPrefix & ".file_number, " & BankDbPrefix & ".qbo_file_name, " & BankDbPrefix & ".ending_balance " _
       & "FROM " & BankDbPrefix & ";"
       
Debug.Print ("Step 4c) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql)
If rst.RecordCount = 0 Then GoTo Exit_Sub
'  Now loop thru the query results parse out the file_number from the account_designator
Do
  Call ExtractACCTidValues(rst!account_designator, _
                          FileNumber, _
                          AccountDesignator, _
                          EndingBalance, _
                          FileName)
                          
  rst.Edit
  rst!file_number = FileNumber
  rst!account_designator = AccountDesignator
  rst!ending_balance = EndingBalance
  rst!qbo_file_name = FileName
  rst.Update
  
  rst.MoveNext
  If rst.EOF Then GoTo Exit_Sub
Loop

Exit_Sub:
  rst.Close
  Exit Sub

End Sub
                              
Private Sub ExtractACCTidValues(ByVal strInput As String, _
                              ByRef FileNumber As String, _
                              ByRef AccountDesignator As String, _
                              ByRef EndingBalance As String, _
                              ByRef FileName As String)
                              
'   "1~Felecia Ellis~1531.84~Ellis download (6).QBO"

Dim I As Long

I = InStr(1, strInput, "~")
FileNumber = Left(strInput, I - 1)
strInput = Right(strInput, Len(strInput) - I)

I = InStr(1, strInput, "~")
AccountDesignator = Left(strInput, I - 1)
strInput = Right(strInput, Len(strInput) - I)

I = InStr(1, strInput, "~")
EndingBalance = Left(strInput, I - 1)
strInput = Right(strInput, Len(strInput) - I)

I = InStr(1, strInput, "~")
FileName = strInput

End Sub

Public Sub RecalcRunningBalance(ByVal BankDbPrefix As String, _
                                 Optional ByVal pHeader As String = "RecalcRunningBalance")

'In the event that transactions are not in date sequence, then recalculate the running balance.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

Dim LastFileNumber   As Long
Dim BalanceCalc      As Currency

' Sort transactions into decending date sequence.......
strSql = "SELECT union_state_temp_qbo.* " _
       & "FROM union_state_temp_qbo " _
       & "ORDER BY union_state_temp_qbo.file_number, union_state_temp_qbo.date_posted DESC , union_state_temp_qbo.amount, union_state_temp_qbo.unique_transaction_id;"
Debug.Print ("(ReCalc Running Bal) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql)
If rst.RecordCount = 0 Then Exit Sub

LastFileNumber = -1 ' Initial file number to force a file break....

'  Now loop thru the query results and recalculate all transaction ending balances..
Do
  If LastFileNumber <> rst!file_number Then
    BalanceCalc = rst!ending_balance
    LastFileNumber = rst!file_number
  End If
  
  If rst!balance <> BalanceCalc Then
   ' Debug.Print (rst!Id & " " & rst!amount & " " & rst!balance & " " & BalanceCalc)
    rst.Edit
    rst!save_balance = rst!balance
    rst!balance = BalanceCalc
    rst.Update
  End If
  BalanceCalc = BalanceCalc - rst!amount
  

  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
  rst.Close

End Sub
Public Function xtst()
  Call RecalcTransactionRunningBalance
End Function

Public Sub RecalcTransactionRunningBalance(Optional ByVal pHeader As String = "RecalcRunningBalance")

'In the event that transactions are not in date sequence, then recalculate the running balance.

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

Dim LastAccountDesignator   As String:  LastAccountDesignator = "???"
Dim BalanceCalc             As Currency

' Sort transactions into decending date sequence.......
strSql = "SELECT transactions.* FROM Transactions " _
       & "ORDER BY transactions.account_designator, " _
       & "         transactions.posted_date, " _
       & "         transactions.Amount DESC, " _
       & "         transactions.qbo_unique_trans_id DESC, " _
       & "         transactions.ID; "
Debug.Print ("(ReCalc Running Bal) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql)
If rst.RecordCount = 0 Then Exit Sub

'  Now loop thru the query results and recalculate all transaction ending balances..
Do
  If LastAccountDesignator <> rst!account_designator Then
    BalanceCalc = 0
    LastAccountDesignator = rst!account_designator
  End If
  
  BalanceCalc = BalanceCalc + rst!amount
  
  ' Debug.Print (rst!Id & " " & rst!amount & " " & rst!running_balance & " " & BalanceCalc)
  rst.Edit
  ' rst!running_balance = BalanceCalc
  rst!calc_balance = BalanceCalc
  rst!mismatch = ""
  If rst!calc_balance <> rst!running_balance Then rst!mismatch = "not equal"
  rst.Update
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
  rst.Close

End Sub


