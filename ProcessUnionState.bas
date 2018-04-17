Attribute VB_Name = "ProcessUnionState"
Option Compare Database
Option Explicit

' Process All files from a Union State Folder....

'  Highlevel steps of design.....
'1 Select Folder
'2 Build QBO temporary QBO table (with additional fields to be pulled from the Transactions database)
'3 Execute Bank/QBO 2 csv
'4 Import csv
'4a  Make sure that this is Union State Bank records.
'4b  Delete un-needed rows that were created by QBO2CSV
'5 Update additional fields from the database
'5a  Integerity Check that all records are reconciled
'6 Export temporary table back to temporary *.xlsx file
'7 Import *.xlsx back into Transactions database to update data.
'8 Clean-up all temporary tables.

'Issues List
'1) All QBO files are not being picked up in one execution of QBO2CSV (Solved)

Dim strSql         As String
Dim HoldSelItem    As String    ' Combined path and file name.
Dim HoldFileName   As String    ' Only the file name.  Path is not included.
Dim HoldFilePath   As String    ' Only the path name.

Dim bankPrefix     As String

Public Function tst()
' Call ImportFid("G:\My Drive\Joel's Files\qbo files\UnionState_NEWQBO_20180217090022.xlsx", True)
'Call ImportFid("C:\Users\johnr\Desktop\VA Services Design\VA Services Data Repository\Normalized Data Tables\Fiduciary.xlsx", False)
Call transactionIntegrityCheck("#1")
'G:\My Drive\Joel's Files\UnionState RLowery\2-23-2018
'Call etlUnionStateQBO("G:\My Drive\Joel's Files\qbo files\download (16).QBO")
'Call etlUnionStateQBO("G:\My Drive\Joel's Files\UnionState RLowery\2-23-2018\download (16).QBO")
'Call etlUnionStateQBO("C:\Users\johnr\Desktop\VA Services Design\VA Services Data Repository\Veterans\Henson, James C  C19393343\Work\download (1).qbo")
'Call etlUnionStateQBO("G:\My Drive\Joel's Files\2017 Bank Accounts\TestingFolder\ProperSoft Running Balance Issue\JGrayTech Metro.qbo")
'Call etlUnionStateQBO("G:\My Drive\Joel's Files\2017 Bank Accounts\TestingFolder\DuplicateFilesIssue\JGrayTech Metro.qbo")
'Call etlUnionStateQBO("G:\My Drive\Joel's Files\2017 Bank Accounts\TestingFolder\OVERLAPPING qbo\Ellis download (6).QBO")
'Call etlUnionStateQBO("G:\My Drive\Joel's Files\2017 Bank Accounts\TestingFolder\Chase2818_Activity_20180127.QBO")
Call etlUnionStateQBO("G:\My Drive\Joel's Files\2017 Bank Accounts\TestingFolder\ProperSoft Running Balance Issue\QBO2CSV\download.QBO")
   
Call transactionIntegrityCheck("#2")
End Function

Public Function etlUnionStateQBO(Optional ByVal preSelectedInput As String = "")

' ETL (Extract, Transform and Load) for Union State Bank
Dim BankFilePrefix As String: BankFilePrefix = "UnionState"  ' Set value of windows file prefix.
Dim BankDbPrefix As String: BankDbPrefix = "union_state_temp_qbo"  ' Set value to used for Temp QBO data table.

Dim I As Long, J As Long
Dim ErrorMsg  As String
Dim ErrHeader As String:  ErrHeader = "Process " & BankFilePrefix & " routine.............."

'1 Select Folder
   HoldFileName = SelectFileQBO(HoldFileName, HoldFilePath, HoldSelItem, preSelectedInput)
   If HoldFileName = "" Then Exit Function   ' No file was chosen.....
   Debug.Print (HoldSelItem)
   Debug.Print (HoldFileName)
   Debug.Print (HoldFilePath)
   
'2 Build QBO temporary QBO table (with additional fields to be pulled from the Transactions database)
   Call buildQboTable(BankDbPrefix)   '  Build "union_state_temp_qbo" database empty table.
   
'3 Execute Bank/QBO 2 csv
'     Input- All QBO files found in the selected folder are combined.  File is created with Qbo2Csv program.
'     Output = UnionStateqbo.csv
   Dim convertPgmName As String: convertPgmName = "Qbo2CSV"
   Dim FileName       As String: FileName = BankFilePrefix
   Dim fileExtension  As String: fileExtension = "qbo"
   Dim qboFileList()  As Variant   '  This list of QBO files will be filled in by "combineTextFiles" routine.
                                   '  List will be used later to rename files.
   Dim qboBankList()  As Variant   '  This list of QBO bank routings will be filled in by "combineTextFiles" routine.
                                   '  List will be used later to process specific bank(s) files.
   
   Dim combinedPathAndFile  As String
   combinedPathAndFile = combineTextFiles(qboFileList, qboBankList, HoldFilePath, fileExtension, FileName)
   If Not IsArrayAllocated(qboFileList) Then
   'If UBound(qboFileList) = 0 And qboFileList(0) = "" Then
      ErrorMsg = "No QBO files were found in " & HoldFilePath & " directory:" & Chr(13) & _
        HoldFilePath
      Call MsgBox(ErrorMsg, vbOKOnly, ErrHeader)
      Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
      Call cleanUpQboFiles(HoldFilePath, "", BankDbPrefix, BankFilePrefix)
      Exit Function
   End If
   Call CreateAndRunProperSoftBATfile(HoldFilePath, _
                                      combinedPathAndFile, _
                                      convertPgmName, _
                                      FileName, _
                                      fileExtension)
   
'4 Import UnionStateqbo.csv into "union_state_temp_qbo"
    strSql = "UPDATE import_spec1_file_name INNER JOIN import_spec2_worksheet_name " _
           & "ON import_spec1_file_name.ID = import_spec2_worksheet_name.input_file_name_ID " _
           & "SET import_spec1_file_name.input_file_name = """ & BankFilePrefix & "qbo.csv"", " _
           & "import_spec2_worksheet_name.work_sheet_name = """ & BankFilePrefix & "qbo"", " _
           & "import_spec2_worksheet_name.output_table_name = """ & BankDbPrefix & """ " _
           & "WHERE (((import_spec1_file_name.spec_name)=""QBO_Import_from_QBO2CSV""));"
    Debug.Print ("4 Import csv - " & strSql)
    DoCmd.SetWarnings False
    DoCmd.RunSQL (strSql)
    DoCmd.SetWarnings True
    If Not ImportFid(HoldFilePath & "\" & BankFilePrefix & "qbo.csv", False) Then
      ErrorMsg = "INITIAL Import Process of raw QBO file encountered errors." & Chr(13) _
        & Chr(13) & "Import Process will be ABORTED so that imported_table_errors can be checked....."
      Call MsgBox(ErrorMsg, vbOKOnly, ErrHeader)
      Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
      End
    End If
    
'4a  Make sure that this is Union State Bank records - Bank-ID=62203395
   ' Call confirmOnlyExpectedBankFound(BankDbPrefix, BankFilePrefix, 62203395)

'4b  Delete un-needed rows from "union_state_temp_qbo" that were created by QBO2CSV
    strSql = "DELETE " & BankDbPrefix & ".*, " & BankDbPrefix & ".amount, " & BankDbPrefix & ".payee " _
           & "FROM " & BankDbPrefix & " " _
           & "WHERE (((" & BankDbPrefix & ".amount) Is Null Or (" & BankDbPrefix & ".amount)=0) AND ((" & BankDbPrefix & ".payee) Is Null Or (" & BankDbPrefix & ".payee)="" ""));"
    Debug.Print ("Step 4b - " & strSql)
    DoCmd.RunSQL (strSql)
    
'4c  Parse out the file_number from the account_designator field.
    Call ParseOutFileNumber(BankDbPrefix)
    
'5 Update additional fields from the database  va_case_id and account_number (based on alpha_id)
    strSql = "UPDATE " & BankDbPrefix & " INNER JOIN (bank_account INNER JOIN bank ON bank_account.Bank = bank.alpha_id) ON (bank_account.account_designator = " & BankDbPrefix & ".account_designator) AND (" & BankDbPrefix & ".bank_id = bank.routing_number) " _
           & "SET " & BankDbPrefix & ".va_case_id = [bank_account].[va_case_id], " & BankDbPrefix & ".account_number = [bank_account].[account_number];"
    Debug.Print ("Step 5 - " & strSql)
    DoCmd.RunSQL (strSql)
    'Because transactions are not in predictable date sequence, then recalculate the running balance.
    Call RecalcRunningBalance(BankDbPrefix)

    
    
'5a  Integerity Check that all records are reconciled the QBO import file.
    Call integrityCheckQBO(BankDbPrefix)
    
'5b  Integrity Check the "transactions" table to ensure we are starting out clean....
    Call transactionIntegrityCheck("etl" & BankFilePrefix & "QBO - Step 5b - (Before new transactions)")
    
'5c  Make a backup copy of "transactions" table...   "SELECT transactions.* INTO transactions_2018a FROM transactions;"
    Dim transactionsBackupName As String
    transactionsBackupName = "transactions_bkup_" & Format(Now(), "yyyymmddhhmmss")
    DoCmd.CopyObject , transactionsBackupName, acTable, "transactions"

'6 Export table "union_state_temp_qbo" back to temporary "UnionState_NEWQBO_yyyymmddhhmmss.xlsx" file
    Dim holdUnionStateExcelFileName As String
    Call exportToExcel(HoldFilePath, _
                      FileName, _
                      ".xlsx", _
                      BankDbPrefix, _
                      holdUnionStateExcelFileName)

'7 Import "UnionState_NEWQBO_yyyymmddhhmmss.xlsx"  back into "transactions" database to update data.
    I = InStrRev(holdUnionStateExcelFileName, "\")
    Debug.Print (Mid(holdUnionStateExcelFileName, I + 1))
    
    strSql = "UPDATE import_spec1_file_name INNER JOIN import_spec2_worksheet_name " _
           & "ON import_spec1_file_name.ID = import_spec2_worksheet_name.input_file_name_ID " _
           & "SET import_spec1_file_name.input_file_name = """ & Mid(holdUnionStateExcelFileName, I + 1) & """, " _
           & "import_spec2_worksheet_name.work_sheet_name = ""QBO"", " _
           & "import_spec2_worksheet_name.output_table_name = ""transactions"" " _
           & "WHERE (((import_spec1_file_name.spec_name)=""" & BankFilePrefix & "Temp""));"
    Debug.Print ("4 Import csv - " & strSql)
    DoCmd.SetWarnings False
    DoCmd.RunSQL (strSql)
    DoCmd.SetWarnings True
    If Not ImportFid(holdUnionStateExcelFileName, False) Then
      ErrorMsg = "Import Process of QBO into ""transactions"" table encountered errors." & Chr(13) _
        & Chr(13) & "Import Process will be ABORTED so that imported_table_errors can be checked....."
      Call MsgBox(ErrorMsg, vbOKOnly, ErrHeader)
      Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
      End
    End If
    Call transactionIntegrityCheck("etl" & BankFilePrefix & "QBO - Step 7 - (After new transactions added)")

'8 Re-Name input QBO files.......   RenameFileOrDir
   Dim fromFileName As String, toFileName As String
   
   Dim RenameMsgText   As String
   RenameMsgText = "List of QBO files in folder """ & HoldFilePath & """ have been processed:" & Chr(13) & Chr(13)
   For I = LBound(qboFileList) To UBound(qboFileList)
      J = InStrRev(qboFileList(I), "\")
      RenameMsgText = RenameMsgText & Mid(qboFileList(I), J + 1) & Chr(13)
   Next I
   
   Dim MsgResponse  As Long
   MsgResponse = MsgBox(RenameMsgText & Chr(13) & "Do you want to rename these as *.bak?", vbYesNo, ErrHeader)
   
   If MsgResponse = vbYes Then GoTo Rename_Process
   
   ErrorMsg = RenameMsgText & Chr(13) & _
      "Files WILL NOT be renamed." & Chr(13) & "Do you want to RESTORE ""transactions"" table to UNDO any changes?" & Chr(13) & Chr(13) & _
      "-- Yes will UNDO and continue, " & Chr(13) & "-- No will NOT UNDO but will continue with workfile cleanup, " & Chr(13) & "-- Cancel will NOT UNDO but will Stop (Without workfile cleanup)."
   MsgResponse = MsgBox(ErrorMsg, vbYesNoCancel, ErrHeader)
   Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))

   If MsgResponse = vbCancel Then End
   If MsgResponse = vbNo Then GoTo Skip_Rename

UNDO_Process:
  Call DelTbl("transactions")
  DoCmd.Rename "transactions", acTable, transactionsBackupName
  GoTo Skip_Rename
  
Rename_Process:
   For I = LBound(qboFileList) To UBound(qboFileList)
      fromFileName = qboFileList(I)
      toFileName = fromFileName & ".bak"
      Debug.Print (toFileName)
      If Not RenameFileOrDir(fromFileName, toFileName) Then
        ErrorMsg = "Rename process FAILED for: " & Chr(13) & fromFileName & Chr(13) & Chr(13) & _
               "Rename process is aborted...."
        Call MsgBox(ErrorMsg, vbOKOnly, ErrHeader)
        Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
        End
      End If
   Next I
Skip_Rename:
   
'9 Clean-up all temporary tables.
    Call cleanUpQboFiles(HoldFilePath, transactionsBackupName, BankDbPrefix, BankFilePrefix)
   
   
   Debug.Print ("Finished")
End Function
Public Function Clone()
 ' Call Clone_Import_Recipe("UnionStateTemp", "NewName2", "MyInputFile.xlsx")
  Call Clone_Import_Recipe("Fiduciary_Excel", "NewName2", "MyFiduciary.xlsx")
  Debug.Print ("Clone is finished.....")
  
End Function
Public Function RemoveIT()
  Call Remove_Import_Recipe("NewName")
  Call Remove_Import_Recipe("NewName2")
  Debug.Print ("RemoveIT is finished....")
End Function

