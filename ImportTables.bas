Attribute VB_Name = "ImportTables"
Option Compare Database
Option Explicit

' Punch List
' 1) Test single key's vs dual keys
' 2) Test date keys as single and multiple
' 3) Gracefully handle the error when Excel Spreadsheet is open by another user.  Permission Denied...
' 4) Develop support for non-unique keys.  (Done)
' 5) Develop support for file rename on error free completion of import.
' 6) Screens to support update of Spec DB
' 7) Error display report.
' 8) Test *.csv import
' 9) Clean up of work files
' 10) Report of Errors
' 11) Screens to maintain spec's
' 12) Test Date fields as Key's and not keys
' 13) Test senario - some fields being totally left out of specification creating a "gap" in the F fields.
' 14) Don't allow "Indexed - Yes (Duplicates OK)" when "unique_import_key_value" is NO
' 15) Add option to support either CSV or XLXS both for the import, without having to duplicate the spec.  (No good way to handle worksheet names)
' 16) Debug Adding Specifications for *.csv
' 17) Support passing a path:file name into the ImportFID and bypass manual selection
' 18) Add support for UNDO if there is an error.

Dim strSql         As String
Dim ErrorMsg       As String
Dim HoldSelItem    As String    ' Combined path and file name.
Dim HoldFileName   As String    ' Only the file name.  Path is not included.
Dim HoldFilePath   As String    ' Only the path name.

'  These Following array values are populated by GetImportSpecifications and includes only Active Items
Dim numOfSpecs   As Long
Dim Add_Date_Changed_to_Rows  As Boolean
Dim aUnique_Dup_Key_Field As Boolean
Dim aActive_flag_name  As String
Dim aMark_active_flag  As String

Dim aOutput_Table_Name(1 To 150) As String
Dim aExcel_Column_Number(1 To 150) As String
Dim aExcel_Heading_Text(1 To 150) As String
Dim aField_Name_Output(1 To 150) As String
Dim aData_Type(1 To 150) As String
Dim aDup_Key_Field(1 To 150) As Boolean
Dim aReject_Err_File(1 To 150) As Boolean
Dim aReject_Err_Rows(1 To 150) As Boolean
Dim aAccept_Changes(1 To 150) As Boolean
Dim aAllowChange2Blank(1 To 150) As Boolean
Dim aSave_Date_Changed(1 To 150) As Boolean

Dim aold_Val(1 To 150) As String      ' Hold the old value from the target table.
Dim afield_err(1 To 150) As String    ' Any error that is found with this field.
Dim holdMergedData(1 To 150) As String
Dim holdKeyNames(1 To 150) As String

'  Variables to hold field values for Specification Populate process.
Dim pID3(1 To 150) As Long
Dim pExcel_heading_text(1 To 150) As String
Dim pExcel_column_number(1 To 150) As String

Dim aWorkSheetNames_len As Long ' For an xlsx file, this array will hold list of tabs.
Dim aWorkSheetNames(1 To 150)  As String
Dim aTablesFound(1 To 150)     As String
Dim aUniqueImportKeyValue(1 To 150) As Boolean
Dim aFileID    As Long  ' Used for import of specs outline from a spreadsheet.
Dim aFileID2   As Long  ' Used for import of specs outline from a spreadsheet.

Dim aPrintDebugLog As Boolean ' If this is True, then the Debug log is produced.
Dim aNoImportErrorsWereFound As Boolean  ' This will be false if there were errors found.
Dim aWorkSheetErrorFound As Boolean      ' This flag will be used to flag an error in a worksheet.

'Dim aTablesFoundCount    As Long  ' This is a list of tables found in the worksheet

Public Function PopulateSpecification(Optional ByVal printDebugLog As Boolean = True)

aPrintDebugLog = printDebugLog ' set the Public flag....
If Not aPrintDebugLog Then Debug.Print ("Debug.Print log for ""ImportFid"" is turned off.....")

HoldFileName = SelectFileP
If HoldFileName = "" Then Exit Function   ' No file was chosen.....

' Now process the selected file along with all Worksheets found in the Import Specification...
Dim jj As Long
   For jj = 1 To aWorkSheetNames_len
     'DebugPrint (aWorkSheetNames(JJ))
     Call ImportWorkSheetSpecs(aWorkSheetNames(jj))
   Next jj

End Function

Private Function ImportWorkSheetSpecs(WksheetName As String)
' This function will handle importing specs for a specific worksheet.

Dim rst   As Recordset
Dim I As Long, J As Long
Dim strSql   As String

Dim Response As Variant

strSql = "SELECT import_spec2_worksheet_name.ID2, import_spec2_worksheet_name.input_file_name_ID, import_spec2_worksheet_name.work_sheet_name, import_spec2_worksheet_name.output_table_name " _
       & "FROM import_spec2_worksheet_name " _
       & "WHERE (((import_spec2_worksheet_name.input_file_name_ID)=" & aFileID _
       & ") AND ((import_spec2_worksheet_name.work_sheet_name)='" & WksheetName & "'));"
DebugPrint ("ImportWorkSheetSpecs - " & strSql)
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
'  Now loop thru the query results and populate the array table
For I = 1 To rst.RecordCount
  ErrorMsg = " Is Target Table-'" & Nz(rst!Output_Table_Name, "") & "' the correct output table?"
  Response = MsgBox(ErrorMsg, vbYesNo, " Verify the table to be used as the target data import.")
  DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
  If Response = vbYes Then Exit For
  rst.MoveNext
  If rst.EOF Then Exit For
Next I

Dim holdID2    As Long
If rst.RecordCount = 0 Or Response = vbNo Then
  rst.Close
  aFileID2 = 0
Else
  aFileID2 = rst!ID2
End If

' At this point, if aFileID2 = 0, this means that we need a new record appended to the table.
If aFileID2 = 0 Then aFileID2 = Append_New_Record(WksheetName)

' At this point, aFileID2 will have the ID2 value, needed for the releted field records.
' Process should first delete all of the related field records and then add them back based on contents of the
' Excel worksheet....

Call Populate_pExcel_Heading_Text(WksheetName)  ' Populate all spec variables from the spreadsheet.
Call Populate_Merge_Field_Specs                 ' Read the field spec table (aFileID2 values) and Merge existing values.




End Function
Private Sub Populate_Merge_Field_Specs()

Dim I As Long, J As Long
Dim strSql As String
Dim rst    As Recordset

strSql = "SELECT import_spec3_fields.* FROM import_spec3_fields WHERE (((import_spec3_fields.worksheet_name_id)=" _
      & aFileID2 & "));"

Set rst = Application.CurrentDb.OpenRecordset(strSql)  ' Open recordset with intent to edit.
If rst.RecordCount = 0 Then GoTo Finished_Do_Loop
'  Now loop thru the query results and populate the array table
Do
  J = 0
  For I = 1 To numOfSpecs  ' Look for existing item in the table.
    If pExcel_heading_text(I) = rst!excel_heading_text Then J = I
    If J <> 0 Then Exit For
  Next I
  If J <> 0 Then  ' If item was found then merge into the table.
    rst.Edit
    pID3(J) = rst!ID3
    rst!excel_column_number = pExcel_column_number(J)  ' Update the column number.
    rst.Update
  End If
  If J = 0 Then  ' If the item was not found then mark the item as inactive.
    rst.Edit
    rst!active3 = False
    rst.Update
  End If
    
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:

'  Now, we need to append the new rows.......
For I = 1 To numOfSpecs
  If pID3(I) = 0 Then
    rst.AddNew
    rst!worksheet_name_id = aFileID2
    rst!key_field = False
    rst!excel_column_number = pExcel_column_number(I)
    rst!excel_heading_text = pExcel_heading_text(I)
    rst!field_name_output = ""
    rst!active3 = True
    rst!accept_changes3 = "Empty"
    rst!allow_change_2_blank3 = "Empty"
    rst!Reject_Err_File3 = "Empty"
    rst!Reject_Err_Rows3 = "Empty"
    rst!save_date_changed3 = "Empty"
    rst.Update
  End If
Next I
rst.Close

End Sub



Private Function Populate_pExcel_Heading_Text(wrkSheetName As String)

Dim ExcelApp As Object
Dim ExcelBook As Object
Dim ExcelSheet As Object

Dim ColumnNumber As String, I As Long
Dim excelColumns() As String  '  Excel column letters array......
Call Init_Excel_Names(excelColumns)

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(HoldSelItem)

If Right(HoldSelItem, 4) = ".csv" Then
   Set ExcelSheet = ExcelBook.Sheets(1)
Else
   Set ExcelSheet = ExcelBook.Sheets(wrkSheetName)
End If

ExcelSheet.Activate

'  Capture the Headings from the Worksheet.
numOfSpecs = 0
For I = 1 To UBound(excelColumns)
  ColumnNumber = excelColumns(I) & "1"  ' Construct the cell id for proper row 1 heading....
  If ExcelSheet.Range(ColumnNumber).Value = "" Then Exit For   ' First blank description will stop capturing.
 'DebugPrint (columnNumber & "   '" & ExcelSheet.Range(columnNumber).Value & "'")
  numOfSpecs = numOfSpecs + 1
  pID3(numOfSpecs) = 0
  pExcel_heading_text(numOfSpecs) = ExcelSheet.Range(ColumnNumber).Value
  pExcel_column_number(numOfSpecs) = excelColumns(I)
Next I

ExcelBook.Saved = True  ' Avoid the user message when closing Excel Workbook
ExcelBook.Close
Set ExcelBook = Nothing
Set ExcelApp = Nothing
Set ExcelSheet = Nothing

End Function

Private Function Append_New_Record(WksheetName As String)

Dim import_spec2_worksheet_name  As dao.Recordset

Set import_spec2_worksheet_name = CurrentDb.OpenRecordset("import_spec2_worksheet_name", dbOpenDynaset)

import_spec2_worksheet_name.AddNew
import_spec2_worksheet_name![input_file_name_ID] = aFileID
import_spec2_worksheet_name![Work_Sheet_Name] = WksheetName
import_spec2_worksheet_name![Output_Table_Name] = ""
import_spec2_worksheet_name![active2] = True
import_spec2_worksheet_name![accept_changes2] = "Empty"
import_spec2_worksheet_name![allow_change_2_blank2] = "Empty"
import_spec2_worksheet_name![Reject_Err_File2] = "Empty"
import_spec2_worksheet_name![Reject_Err_Rows2] = "Empty"
import_spec2_worksheet_name![save_date_changed2] = "Empty"
import_spec2_worksheet_name.Update

import_spec2_worksheet_name.Bookmark = import_spec2_worksheet_name.LastModified
'DebugPrint ("import_spec2_worksheet_name.Bookmark=" & import_spec2_worksheet_name.Bookmark)
import_spec2_worksheet_name.MovePrevious
import_spec2_worksheet_name.MoveNext
aFileID2 = import_spec2_worksheet_name!ID2
'DebugPrint ("import_spec2_worksheet_name!ID2=" & import_spec2_worksheet_name!ID2)
import_spec2_worksheet_name.Close
Set import_spec2_worksheet_name = Nothing

End Function

Private Function FilePatternID()

Dim rst As dao.Recordset
Dim matchCount     As Long
Dim matchedName(1 To 150) As String
Dim matchedID(1 To 150) As String
Dim matchedDisplayList  As String
Dim Response As Variant
Dim I  As Long
Dim ErrorMsg  As String

strSql = "SELECT import_spec1_file_name.ID, import_spec1_file_name.input_file_name, import_spec1_file_name.active FROM import_spec1_file_name;"
strSql = "SELECT import_spec1_file_name.* FROM import_spec1_file_name;"
   
FilePatternID = 0
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
   rst.Close
   Exit Function
End If
'  Now loop thru the query results and populate the array table
Do
  If HoldFileName Like rst!Input_File_Name Then
    If matchCount = UBound(matchedName) Then
      ErrorMsg = " matchedName table has been overrun.  Too many MATCHING file entries in import_spec1_file_name table " & Chr(13) & _
              " for the selected file-" & HoldFileName
      MsgBox (ErrorMsg)
      Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
      End
    End If
    matchCount = matchCount + 1
    matchedName(matchCount) = rst!Input_File_Name
    matchedID(matchCount) = rst!Id
    FilePatternID = rst!Id
   'DebugPrint (rst!Id & "  " & rst!Input_File_Name)
  End If
    
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
If matchCount = 0 Then Exit Function

If matchCount > 1 Then
  ErrorMsg = " Multiple matches in the import_spec1_file_name table." & Chr(13) & Chr(13) & _
                    " Do you want to use the first match-" & matchedName(1) & " ?"
  Response = MsgBox(ErrorMsg, vbYesNoCancel)
  DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
  I = 1
  Do While (Response = vbNo)
    If I = matchCount Then Exit Do

    I = I + 1
    ErrorMsg = " Multiple matches in the import_spec1_file_name table." & Chr(13) & Chr(13) & _
                    " Do you want to use the next match-" & matchedName(I) & " ?"
  Response = MsgBox(ErrorMsg, vbYesNoCancel)
  DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
  Loop
End If
rst.Close

If Response = vbYes Then
  FilePatternID = matchedID(I)
  Exit Function
End If

If Response = vbNo Then
  MsgBox (" Action was cancelled. Process will not complete.... ")
  DebugPrint (" Action was cancelled. Process will not complete.... ")
  FilePatternID = 0
End If

If Response = vbCancel Then
  MsgBox (" Action was cancelled. Process will not complete.... ")
  DebugPrint (" Action was cancelled. Process will not complete.... ")
  FilePatternID = 0
End If

End Function

Private Function SelectFileP()   '  Selecte File for Populate
   'Dim fd As Office.FileDialog
   'Set fd = Application.FileDialog(msoFileDialogFilePicker)
   Dim fd As Object
   Dim pos   As Long
   Dim Response  '  Yes=6  No=7  Retry=4  OK=1  Cancel=2
   
Retry_Selection:
   '  Set Initial returning values
   HoldSelItem = ""

   
   Set fd = Application.FileDialog(1)
   If Nz(HoldFilePath, "") = "" Then HoldFilePath = CurrentProject.Path
   With fd
      .InitialFileName = "" & HoldFilePath & "\*.xlsx"
      .Title = "Select a File"
      .Filters.Clear
      .Filters.Add "Excel Files", "*.xlsx"
      If .Show Then HoldSelItem = .SelectedItems(1)
      SelectFileP = HoldSelItem   '  This is the whole path and file name.
      
      If HoldSelItem = "" Then Exit Function
     '  Parse out the name of the file....
      pos = InStrRev(HoldSelItem, "\")
      If pos <> 0 And pos <> Len(HoldSelItem) And pos <> 1 Then
         SelectFileP = Mid(HoldSelItem, pos + 1)
         HoldFileName = Mid(HoldSelItem, pos + 1)
         HoldFilePath = Mid(HoldSelItem, 1, pos - 1)
      End If
      
   End With
   
   If Right(SelectFileP, 5) <> ".xlsx" And _
      Right(SelectFileP, 4) <> ".csv" And _
      Right(SelectFileP, 4) <> ".xls" Then
      
      ErrorMsg = "INVALID file-'" & SelectFileP & "' was chosen." & Chr(13) & Chr(13) _
        & "Must be *.xlsx, *.csv, or *.xls to be valid." & Chr(13) & Chr(13) _
        & "Retry to select another file,  Cancel to Quit."
      Response = MsgBox(ErrorMsg, vbRetryCancel)
      DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
      If Response = 4 Then  ' Retry Button
        GoTo Retry_Selection
        Else
          SelectFileP = ""
          HoldFileName = ""
          Set fd = Nothing
          Exit Function
      End If
      Exit Function
   End If
   
   Set fd = Nothing
   
   aFileID = FilePatternID
   If aFileID = 0 Then Exit Function
   
  'DebugPrint (" FilePatternID=" & aFileID) ' Make sure that filename has already been entered for this file.
   Call GetExcelWrksheetsP(HoldSelItem)      ' GetExcelWrksheets will populate the aWorkSheetNames array....
End Function


Private Sub GetExcelWrksheetsP(excelname As String)
'  aWorkSheetNames is the main output array of tab names.
Dim AppExcel As New Excel.Application, Wkb As Workbook, Wksh As Worksheet
Dim obj As AccessObject, dbs As Object, tempTable As String, spaceIn As Long
Dim I As Long
Dim displayList   As String


If Right(HoldFileName, 4) = ".csv" Then
  aWorkSheetNames_len = aWorkSheetNames_len + 1
  aWorkSheetNames(aWorkSheetNames_len) = Left(HoldFileName, Len(HoldFileName) - 4)
  Exit Sub
End If

On Error GoTo Errorcatch

Dim Response    As Variant

GetListOfWorksheets:
displayList = ""
aWorkSheetNames_len = 0    '  Reset array dimensions....
Set Wkb = AppExcel.Workbooks.Open(excelname)
For Each Wksh In Wkb.Worksheets

  ErrorMsg = " Should Worksheet-" & Wksh.Name & " be processed? "
  Response = MsgBox(ErrorMsg, vbYesNoCancel, _
            "  Checking for Worksheets found in this Workbook-" & HoldFileName & "  ")
  DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
          
  If Response = vbCancel Then Exit For

  If Response = vbYes Then
    aWorkSheetNames_len = aWorkSheetNames_len + 1
    aWorkSheetNames(aWorkSheetNames_len) = Wksh.Name
    displayList = displayList & Wksh.Name & Chr(13)
  End If
Next Wksh

Wkb.Saved = True
Wkb.Close

ErrorMsg = " Are these the Worksheet names for building field import specifications? " & Chr(13) & Chr(13) & displayList
Response = MsgBox(ErrorMsg, _
                  vbYesNoCancel, _
                  "  Checking for Worksheets found in this Workbook-" & HoldFileName & "  ")
DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
If Response = vbNo Then GoTo GetListOfWorksheets
If Response = vbCancel Then aWorkSheetNames_len = 0

AppExcel.Quit
Set Wkb = Nothing
Set AppExcel = Nothing
Exit Sub

Errorcatch:
MsgBox err.Description
If IsNull(Wkb) = False Then Exit Sub
Wkb.Saved = True
Wkb.Close
AppExcel.Quit
Set Wkb = Nothing
Set AppExcel = Nothing

End Sub



   
Public Function ImportFid(Optional ByVal preSelectedInput As String = "", _
                          Optional ByVal printDebugLog As Boolean = True, _
                          Optional ByVal pHeader4Msgs As String) As Boolean
   
' This function will return True is import has NO errors.  False if import finished with errors.
   ImportFid = True  ' Set initial value, assuming no errors will be found.
   
   Dim I As Long
   Dim strSql  As String
   
   aPrintDebugLog = printDebugLog ' set the Public flag....
   If Not aPrintDebugLog Then DebugPrint ("Debug.Print log for ""ImportFid"" is turned off.....")

   HoldFileName = SelectFile(preSelectedInput)
   If HoldFileName = "" Then
    DebugPrint ("File Selection was canceled by the user.  ImportFid will end...")
    Exit Function   ' No file was chosen.....
   End If
   DebugPrint (HoldSelItem)
   
   aNoImportErrorsWereFound = True ' Initial value of this flag
   Call DelTbl("imported_table_errors")
   
   ' Now process the selected file along with all Worksheets found in the Import Specification...
   Dim jj As Long
  'DebugPrint (HoldFileName & " List of valid Excel tabs for " & HoldFileName & "................")
   For jj = 1 To aWorkSheetNames_len
     aWorkSheetErrorFound = False '  Clear this flag before processing a worksheet.
     Call Process_A_Worksheet(aWorkSheetNames(jj), aTablesFound(jj))
     DelTbl ("imported_table_" & aWorkSheetNames(jj)) ' Clean up previous table.
     If Not aWorkSheetErrorFound Then
       DelTbl ("imported_table")
      Else
        DoCmd.Rename "imported_table_" & aWorkSheetNames(jj), acTable, "imported_table"
      End If
   Next jj
      
   ImportFid = aNoImportErrorsWereFound '  Set the return value...
      
   If preSelectedInput = "" Then MsgBox ("HoldFileName=" & HoldFileName)
   Call DelTbl("imported_table_count")

   Exit Function
   
   

End Function

Private Function Selected_XLS_File_Is_Good()
Dim I As Long, J As Long
Dim numOfSpecItems   As Long

Selected_XLS_File_Is_Good = False
    aWorkSheetNames_len = 0
Dim testvariable As String
      ' We need to see if there are any matching records in the specification.
      ' If there are none then select another file.
      ' Next neet to verify that the headings all match.
      ' If one does not match, then select another file.

Call GetExcelWrksheets(HoldSelItem)   ' GetExcelWrksheets will populate the aWorkSheetNames array....
If aWorkSheetNames_len <> 0 Then Selected_XLS_File_Is_Good = True

Exit Function
End Function

Private Sub GetExcelWrksheets(excelname As String)
'  aWorkSheetNames is the main output array of tab names.
Dim AppExcel As New Excel.Application, Wkb As Workbook, Wksh As Worksheet
Dim obj As AccessObject, dbs As Object, tempTable As String, spaceIn As Long
Dim I As Long

If Right(HoldFileName, 4) = ".csv" Then
   Call GetImportSpecifications(Left(HoldFileName, Len(HoldFileName) - 4), "*")
   Exit Sub
End If

'On Error GoTo Errorcatch

aWorkSheetNames_len = 0    '  Reset array dimensions....
Set Wkb = AppExcel.Workbooks.Open(excelname)
For Each Wksh In Wkb.Worksheets
   Call GetImportSpecifications(Wksh.Name, "*")
Next Wksh

Wkb.Saved = True
Wkb.Close

DebugPrint ("aWorkSheetNames_len=" & aWorkSheetNames_len)
For I = 1 To aWorkSheetNames_len
  DebugPrint (aWorkSheetNames(I) & " " & aTablesFound(I))
Next I


AppExcel.Quit
Set Wkb = Nothing
Set AppExcel = Nothing
Exit Sub

Errorcatch:
MsgBox err.Description
'If IsNull(Wkb) = False Then Exit Sub
Wkb.Saved = True
Wkb.Close
AppExcel.Quit
Set Wkb = Nothing
Set AppExcel = Nothing

End Sub

Private Function Add_Worksheet_to_aWorkSheetNames(WksheetName As String, tableName As String, unique_import_key_value As Boolean)
   
Dim I As Long, J As Long

'  Check to see if entry already exists.
For J = 1 To aWorkSheetNames_len
  If aWorkSheetNames(J) = WksheetName And aTablesFound(J) = tableName Then Exit Function
Next J

'  Entry did not exist, add to the table.
If aWorkSheetNames_len = UBound(aWorkSheetNames) Then
  ErrorMsg = " File-" & HoldFileName & Chr(13) & _
           " Worksheets table has been overrun with more than " & aWorkSheetNames_len & " entries.  " & Chr(13) & _
           " Tables will need to be expanded for the file to be processed. "
  MsgBox (ErrorMsg)
  DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
  Exit Function
End If
aWorkSheetNames_len = aWorkSheetNames_len + 1
I = aWorkSheetNames_len
aWorkSheetNames(I) = WksheetName
aTablesFound(I) = tableName
aUniqueImportKeyValue(I) = unique_import_key_value

DebugPrint ("aWorkSheetNames_len=" & aWorkSheetNames_len)
For I = 1 To aWorkSheetNames_len
  DebugPrint (aWorkSheetNames(I) & " " & aTablesFound(I))
Next I

End Function

Private Function SelectFile(Optional ByVal preSelectedFile As String = "")
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
      
      SelectFile = HoldSelItem   '  This is the whole path and file name.
      
      If HoldSelItem = "" Then Exit Function
     '  Parse out the name of the file....
      pos = InStrRev(HoldSelItem, "\")
      If pos <> 0 And pos <> Len(HoldSelItem) And pos <> 1 Then
         SelectFile = Mid(HoldSelItem, pos + 1)
         HoldFileName = Mid(HoldSelItem, pos + 1)
         HoldFilePath = Mid(HoldSelItem, 1, pos - 1)
      End If
      
   End With
   
   If Right(SelectFile, 5) <> ".xlsx" And _
      Right(SelectFile, 4) <> ".csv" And _
      Right(SelectFile, 4) <> ".xls" Then
      
      ErrorMsg = "INVALID file-'" & SelectFile & "' was chosen." & Chr(13) & Chr(13) _
        & "Must be *.xlsx, *.csv, or *.xls to be valid." & Chr(13) & Chr(13) _
        & "Retry to select another file,  Cancel to Quit."
      Response = MsgBox(ErrorMsg, vbRetryCancel)
      DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
      If Response = 4 Then  ' Retry Button
        GoTo Retry_Selection
        Else
          SelectFile = ""
          HoldFileName = ""
          Set fd = Nothing
          Exit Function
      End If
      Exit Function
   End If
   
   If Not Selected_XLS_File_Is_Good Then
      ErrorMsg = "File-'" & SelectFile & "' was chosen." & Chr(13) & Chr(13) _
        & "No specifications for this file were found in the 'import_spec' table." & Chr(13) & Chr(13) _
        & "Retry to select another file,  Cancel to Quit."
      Response = MsgBox(ErrorMsg, vbRetryCancel)
      DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
      If Response = 4 Then  ' Retry Button
        GoTo Retry_Selection
        Else
          SelectFile = ""
          HoldFileName = ""
          Set fd = Nothing
          Exit Function
      End If
   End If
   
   ' Now check to see if headings on the Excel worksheet
   Dim jj As Long
   Dim parmErr As String
   For jj = 1 To aWorkSheetNames_len
     If aWorkSheetNames(jj) <> "" Then
        ErrorMsg = Headings_Mismatch_Error(aWorkSheetNames(jj))
        If ErrorMsg <> "" Then
          ErrorMsg = " File-'" & SelectFile & "' was chosen." & Chr(13) & Chr(13) _
            & " Excel Headings for Workbook " & aWorkSheetNames(jj) & " do not match the specification." & Chr(13) & Chr(13) _
            & ErrorMsg & Chr(13) & "Retry to select another file,  Cancel to Quit."
          Response = MsgBox(ErrorMsg, vbRetryCancel)
          DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
          If Response = 4 Then  ' Retry Button
            GoTo Retry_Selection
          Else
            SelectFile = ""
            HoldFileName = ""
            Set fd = Nothing
            Exit Function
          End If
        
        End If
      End If
   
   Next

   Set fd = Nothing
End Function

Private Function Headings_Mismatch_Error(wrkSheetName As String)

Dim ExcelApp As Object
Dim ExcelBook As Object
Dim ExcelSheet As Object
Dim I As Long, J As Long, ColumnNumber

Dim excelColumns() As String  '  Excel column letters array......
Call Init_Excel_Names(excelColumns)
   
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(HoldSelItem)

If Right(HoldFileName, 4) = ".csv" Then
   Set ExcelSheet = ExcelBook.Sheets(1)
Else
   Set ExcelSheet = ExcelBook.Sheets(wrkSheetName)
End If

Call GetImportSpecifications(wrkSheetName, "*") ' Populate spec variables...

ExcelSheet.Activate
For I = 1 To numOfSpecs
   ColumnNumber = aExcel_Column_Number(I) & "1"  ' Construct the cell id for proper row 1 heading....
   If aExcel_Heading_Text(I) <> ExcelSheet.Range(ColumnNumber).Value Then
     Headings_Mismatch_Error = Headings_Mismatch_Error & _
                               "Cell-" & ColumnNumber & "  Expected-'" & aExcel_Heading_Text(I) & "'  but Found-'" & _
                               ExcelSheet.Range(ColumnNumber).Value & "'" & Chr(13)
   End If
Next I

leaveFunction:
    ExcelBook.Saved = True  ' Avoid the user message when closing Excel Workbook
    ExcelBook.Close
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    Set ExcelSheet = Nothing
    Exit Function

Dim jj As Long
DebugPrint ("HoldFileName-" & HoldFileName & "  WrkSheetName-" & wrkSheetName & "  ................................")
For jj = 1 To numOfSpecs
  DebugPrint (aOutput_Table_Name(jj) & "/" & _
               aExcel_Column_Number(jj) & "/" & _
               aExcel_Heading_Text(jj) & "/" & _
               aField_Name_Output(jj) & "/" & _
               aData_Type(jj) & "/" & _
               aDup_Key_Field(jj) & "/" & _
               aReject_Err_File(jj) & "/" & _
               aReject_Err_Rows(jj) & "/" & _
               aAccept_Changes(jj) & "/" & _
               aAllowChange2Blank(jj) & "/" & _
               aSave_Date_Changed(jj))


Next jj

End Function


Private Function Get_Row_Number_Column(wrkSheetName As String)

Dim ExcelApp As Object
Dim ExcelBook As Object
Dim ExcelSheet As Object
Dim I As Long, J As Long, ColumnNumber

Dim excelColumns() As String  '  Excel column letters array......
Call Init_Excel_Names(excelColumns)
   
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(HoldSelItem)

If Right(HoldFileName, 4) = ".csv" Then
   Set ExcelSheet = ExcelBook.Sheets(1)
Else
   Set ExcelSheet = ExcelBook.Sheets(wrkSheetName)
End If

Call GetImportSpecifications(wrkSheetName, "*") ' Populate spec variables...

ExcelSheet.Activate
For I = 1 To numOfSpecs
   ColumnNumber = aExcel_Column_Number(I) & "1"  ' Construct the cell id for proper row 1 heading....
   If aExcel_Heading_Text(I) <> ExcelSheet.Range(ColumnNumber).Value Then
     Headings_Mismatch_Error = Headings_Mismatch_Error & _
                               "Cell-" & ColumnNumber & "  Expected-'" & aExcel_Heading_Text(I) & "'  but Found-'" & _
                               ExcelSheet.Range(ColumnNumber).Value & "'" & Chr(13)
   End If
Next I

leaveFunction:
    ExcelBook.Saved = True  ' Avoid the user message when closing Excel Workbook
    ExcelBook.Close
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    Set ExcelSheet = Nothing
    Exit Function

Dim jj As Long
DebugPrint ("HoldFileName-" & HoldFileName & "  WrkSheetName-" & wrkSheetName & "  ................................")
For jj = 1 To numOfSpecs
  DebugPrint (aOutput_Table_Name(jj) & "/" & _
               aExcel_Column_Number(jj) & "/" & _
               aExcel_Heading_Text(jj) & "/" & _
               aField_Name_Output(jj) & "/" & _
               aData_Type(jj) & "/" & _
               aDup_Key_Field(jj) & "/" & _
               aReject_Err_File(jj) & "/" & _
               aReject_Err_Rows(jj) & "/" & _
               aAccept_Changes(jj) & "/" & _
               aAllowChange2Blank(jj) & "/" & _
               aSave_Date_Changed(jj))


Next jj

End Function


Private Sub GetImportSpecifications(FindWorkSheetName As String, _
                            FindTableName As String)
                         
Dim rst As dao.Recordset
Dim strSql   As String
Dim I   As Long
Dim holdR   As String


numOfSpecs = 0 ' Initial value / Clear the output arrays
strSql = "SELECT " _
       & "import_spec2_worksheet_name.unique_import_key_value, " _
       & "import_spec2_worksheet_name.active_flag_name, " _
       & "import_spec2_worksheet_name.mark_active_flag, " _
       & "import_spec1_file_name.input_file_name, import_spec2_worksheet_name.work_sheet_name, import_spec2_worksheet_name.output_table_name, " _
       & "import_spec3_fields.key_field, import_spec3_fields.excel_column_number, import_spec3_fields.excel_heading_text, import_spec3_fields.field_name_output, import_spec3_fields.Key_Field, " _
       & "import_spec1_file_name.accept_changes, import_spec2_worksheet_name.accept_changes2, import_spec3_fields.accept_changes3, " _
       & "import_spec1_file_name.allow_change_2_blank, import_spec2_worksheet_name.allow_change_2_blank2, import_spec3_fields.allow_change_2_blank3, " _
       & "import_spec1_file_name.reject_err_file, import_spec2_worksheet_name.reject_err_file2, import_spec3_fields.reject_err_file3, " _
       & "import_spec1_file_name.reject_err_rows, import_spec2_worksheet_name.reject_err_rows2, import_spec3_fields.reject_err_rows3, " _
       & "import_spec1_file_name.save_date_changed, import_spec2_worksheet_name.save_date_changed2, import_spec3_fields.save_date_changed3 " _
       & "FROM (import_spec1_file_name INNER JOIN import_spec2_worksheet_name ON import_spec1_file_name.ID = import_spec2_worksheet_name.input_file_name_ID) INNER JOIN import_spec3_fields ON import_spec2_worksheet_name.ID2 = import_spec3_fields.worksheet_name_id " _
       & "WHERE (((import_spec1_file_name.Active) = True) And ((import_spec2_worksheet_name.active2) = True) And ((import_spec3_fields.active3) = True)) " _
       & "ORDER BY import_spec1_file_name.input_file_name, import_spec2_worksheet_name.work_sheet_name, import_spec2_worksheet_name.output_table_name, import_spec3_fields.excel_column_number;"
DebugPrint ("GetImportSpecifications=" & strSql)
 
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
   rst.Close
   Exit Sub
End If

Do
  If (HoldFileName Like Nz(rst!Input_File_Name, "")) And _
     (Nz(rst!Output_Table_Name) Like FindTableName) And _
     (FindWorkSheetName Like Nz(rst!Work_Sheet_Name, "")) Then
      
      If I = UBound(aOutput_Table_Name) Then
         ErrorMsg = " File-" & HoldFileName & " / " & FindWorkSheetName & Chr(13) & _
                 " Spec Items table has been overrun with more than " & I & " entries.  " & Chr(13) & _
                 " Tables will need to be expanded for the file to be processed. "
         MsgBox (ErrorMsg)
         DebugPrint (Chr(13) & "****" & ErrorMsg & Chr(13))
         Exit Sub
      End If
      
      I = I + 1
      numOfSpecs = I
      aOutput_Table_Name(I) = Nz(rst!Output_Table_Name, "")
      aExcel_Column_Number(I) = Nz(rst!excel_column_number, "")
      aExcel_Heading_Text(I) = Nz(rst!excel_heading_text, "")
      aField_Name_Output(I) = Nz(rst!field_name_output, "")
     
      aData_Type(I) = ""  ' Default value....
      If aOutput_Table_Name(I) <> "" Then _
        aData_Type(I) = Field_Type(aOutput_Table_Name(I), aField_Name_Output(I))
      DebugPrint (1 & " aOutput_Table_Name(i)=" & aOutput_Table_Name(I) & "  " & _
                  "aField_Name_Output(i)=" & aField_Name_Output(I) & "  " & _
                  "aData_Type(i)=" & aData_Type(I))
      
      aDup_Key_Field(I) = Nz(rst!key_field, False)
      aUnique_Dup_Key_Field = Nz(rst!unique_import_key_value, "")
      aActive_flag_name = Nz(rst!active_flag_name, "")
      aMark_active_flag = Nz(rst!mark_active_flag, "")
      
      If FindTableName = "*" Then _
        Call Add_Worksheet_to_aWorkSheetNames(FindWorkSheetName, rst!Output_Table_Name, rst!unique_import_key_value)
      
      aReject_Err_File(I) = False
      holdR = ""
      If Nz(rst!Reject_Err_File, "Empty") <> "Empty" Then holdR = rst!Reject_Err_File
      If Nz(rst!Reject_Err_File2, "Empty") <> "Empty" Then holdR = rst!Reject_Err_File2
      If Nz(rst!Reject_Err_File3, "Empty") <> "Empty" Then holdR = rst!Reject_Err_File3
      If Left(holdR, 1) = "Y" Then aReject_Err_File(I) = True
      
      aReject_Err_Rows(I) = False
      holdR = ""
      If Nz(rst!Reject_Err_Rows, "Empty") <> "Empty" Then holdR = rst!Reject_Err_Rows
      If Nz(rst!Reject_Err_Rows2, "Empty") <> "Empty" Then holdR = rst!Reject_Err_Rows2
      If Nz(rst!Reject_Err_Rows3, "Empty") <> "Empty" Then holdR = rst!Reject_Err_Rows3
      If Left(holdR, 1) = "Y" Then aReject_Err_Rows(I) = True
      
      aAccept_Changes(I) = True
      holdR = ""
      If Nz(rst!Accept_Changes, "Empty") <> "Empty" Then holdR = rst!Accept_Changes
      If Nz(rst!accept_changes2, "Empty") <> "Empty" Then holdR = rst!accept_changes2
      If Nz(rst!accept_changes3, "Empty") <> "Empty" Then holdR = rst!accept_changes3
      If Left(holdR, 1) = "N" Then aAccept_Changes(I) = False
      
      aAllowChange2Blank(I) = False
      holdR = ""
      If Nz(rst!allow_change_2_blank, "Empty") <> "Empty" Then holdR = rst!allow_change_2_blank
      If Nz(rst!allow_change_2_blank2, "Empty") <> "Empty" Then holdR = rst!allow_change_2_blank2
      If Nz(rst!allow_change_2_blank3, "Empty") <> "Empty" Then holdR = rst!allow_change_2_blank3
      If Left(holdR, 1) = "Y" Then aAllowChange2Blank(I) = True
      
      aSave_Date_Changed(I) = False
      holdR = ""
      If Nz(rst!Save_Date_Changed, "Empty") <> "Empty" Then holdR = rst!Save_Date_Changed
      If Nz(rst!save_date_changed2, "Empty") <> "Empty" Then holdR = rst!save_date_changed2
      If Nz(rst!save_date_changed3, "Empty") <> "Empty" Then holdR = rst!save_date_changed3
      If Left(holdR, 1) = "Y" Then aSave_Date_Changed(I) = True
      
      If aSave_Date_Changed(I) Then
        Add_Date_Changed_to_Rows = True
      End If
    End If
    rst.MoveNext
    If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
rst.Close
Set rst = Nothing


End Sub

Private Sub Init_Excel_Names(ColumnNames() As String)

Dim alpha  As String
Dim I As Long, J As Long, jj As Long
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         
Dim U   As Long
On Error Resume Next
ReDim ColumnNames(676)
On Error GoTo 0
U = UBound(ColumnNames)
If U > 702 Then U = 702

For I = 0 To Len(alpha)
   For J = 1 To Len(alpha)
      If jj = UBound(ColumnNames) Then Exit Sub
      jj = jj + 1
      If I <> 0 Then ColumnNames(jj) = Mid(alpha, I, 1) & Mid(alpha, J, 1)
      If I = 0 Then ColumnNames(jj) = Mid(alpha, J, 1)
   Next J
Next I
End Sub



Private Sub printImportSpecs(FileName As String, wrkSheetName As String)

Dim jj As Long
DebugPrint ("Filename-" & FileName & "  WrkSheetName-" & wrkSheetName & "  ................................")
DebugPrint ("aOutput_Table_Name/" & _
             "aExcel_Column_Number/" & _
             "aExcel_Heading_Text/" & _
             "aField_Name_Output/" & _
             "aData_Type/" & _
             "aDup_Key_Field/" & _
             "aReject_Err_File/" & _
             "aReject_Err_Rows/" & _
             "aAccept_Changes/" & _
             "aAllowChange2Blank/" & _
             "aSave_Date_Changed")
For jj = 1 To numOfSpecs
  DebugPrint (aOutput_Table_Name(jj) & "/" & _
               aExcel_Column_Number(jj) & "/" & _
               aExcel_Heading_Text(jj) & "/" & _
               aField_Name_Output(jj) & "/" & _
               aData_Type(jj) & "/" & _
               aDup_Key_Field(jj) & "/" & _
               aReject_Err_File(jj) & "/" & _
               aReject_Err_Rows(jj) & "/" & _
               aAccept_Changes(jj) & "/" & _
               aAllowChange2Blank(jj) & "/" & _
               aSave_Date_Changed(jj))
Next jj
End Sub


Private Function xlColNum(ColumnName As String)
' Input is the Alpha column number, output is the numeric equivilant
Dim alpha  As String
Dim I As Long, J As Long, jj As Long
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

I = InStr(alpha, Mid(ColumnName, 1, 1))
xlColNum = I
If Len(ColumnName) = 2 Then
  J = InStr(alpha, Mid(ColumnName, 2, 1))
  xlColNum = (I * 26) + J
End If
      
End Function

Private Function xlColAlfa(ColumnNum As Long)
' Input to function is numeric excel column, and output is the Alpha equivilant.
Dim I As Long, J As Long, jj As Long, alpha As String
alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

J = ColumnNum \ 26
I = ColumnNum Mod 26
If I = 0 Then
  I = I + 26
  J = J - 1
End If
xlColAlfa = Mid(alpha, I, 1)
If J > 0 Then xlColAlfa = Mid(alpha, J, 1) & xlColAlfa

End Function


Private Function AddTableFields(Output_Table As String)

Dim tdf       As TableDef
Dim target    As TableDef
Dim db As dao.Database, rst As Recordset
Dim fld As dao.Field
Dim prop As dao.Property
Dim I As Long, J As Long, myDataType As Long

Set db = CurrentDb
Set tdf = db.TableDefs("imported_table")
Set target = db.TableDefs(Output_Table)

' Create the new fields in the work table....
For I = 1 To numOfSpecs   '  Add new fields to the Work Table  "imported_table"
   If aField_Name_Output(I) <> "" Then
     myDataType = 10   '  Text
     If aData_Type(I) = "date" Then myDataType = 8   '  dbDate  8
     If aData_Type(I) = "number" Then myDataType = 7 '  dbDouble 7
     tdf.Fields.Append tdf.CreateField(aField_Name_Output(I), myDataType)
     tdf.Fields.Append tdf.CreateField(aField_Name_Output(I) & "_err", 10)
   End If
Next I
tdf.Fields.Append tdf.CreateField("ErrorText", 10)     ' dbText
tdf.Fields.Append tdf.CreateField("matched_target_table", 10)     ' dbText
tdf.Fields.Append tdf.CreateField("row_contains_error", 10)       ' dbText
tdf.Fields.Append tdf.CreateField("reject_err_row", 10)       ' dbText
tdf.Fields.Append tdf.CreateField("row_has_changed", 10)       ' dbText

If Add_Date_Changed_to_Rows Then
  tdf.Fields.Append tdf.CreateField("create_date_time", 8)     ' dbDate     Now()
  tdf.Fields.Append tdf.CreateField("create_user", 10)     ' dbText         HoldSelItem
  tdf.Fields.Append tdf.CreateField("create_program", 10)     ' dbText      "ImportFID"
  tdf.Fields.Append tdf.CreateField("create_file", 10)     ' dbText         HoldSelItem
  
  tdf.Fields.Append tdf.CreateField("update_date_time", 8)     ' dbDate     Now()
  tdf.Fields.Append tdf.CreateField("update_user", 10)     ' dbText         HoldSelItem
  tdf.Fields.Append tdf.CreateField("update_program", 10)     ' dbText      "ImportFID"
  tdf.Fields.Append tdf.CreateField("update_file", 10)     ' dbText         HoldSelItem
  
  If Not FieldExists("create_date_time", Output_Table) Then target.Fields.Append target.CreateField("create_date_time", 8)  ' dbDate     Now()
  If Not FieldExists("create_user", Output_Table) Then target.Fields.Append target.CreateField("create_user", 10)     ' dbText         HoldSelItem
  If Not FieldExists("create_program", Output_Table) Then target.Fields.Append target.CreateField("create_program", 10)     ' dbText      "ImportFID"
  If Not FieldExists("create_file", Output_Table) Then target.Fields.Append target.CreateField("create_file", 10)     ' dbText         HoldSelItem
  
  If Not FieldExists("update_date_time", Output_Table) Then target.Fields.Append target.CreateField("update_date_time", 8)  ' dbDate     Now()
  If Not FieldExists("update_user", Output_Table) Then target.Fields.Append target.CreateField("update_user", 10)     ' dbText         HoldSelItem
  If Not FieldExists("update_program", Output_Table) Then target.Fields.Append target.CreateField("update_program", 10)     ' dbText      "ImportFID"
  If Not FieldExists("update_file", Output_Table) Then target.Fields.Append target.CreateField("update_file", 10)     ' dbText         HoldSelItem
End If

' Now add any Temporary fields to the traget table.
If FieldExists("excel_row_number", Output_Table) Then target.Fields.Delete ("excel_row_number")
target.Fields.Append target.CreateField("excel_row_number", 7)     '  dbDouble

Dim srcField   As String   ' Source Field string......
strSql = ""
For I = 1 To numOfSpecs '  Update data fields in the imported_table....
  If aField_Name_Output(I) <> "" Then
   If srcField = "" Then strSql = "UPDATE imported_table SET "
   srcField = "[imported_table].[F" & xlColNum(aExcel_Column_Number(I)) & "]"
   If aData_Type(I) = "date" Then srcField = "CVDate(" & srcField & ")"
   srcField = "imported_table.[" & aField_Name_Output(I) & "] = " & srcField & ", "
   strSql = strSql + srcField
  End If
Next I
' Include Update Stamp
If Add_Date_Changed_to_Rows Then
  strSql = strSql & "imported_table.create_date_time = Now(), " & _
                    "imported_table.create_user = " & Scrub(Environ("USERNAME")) & ", " & _
                    "imported_table.create_program = 'ImportFID', " & _
                    "imported_table.create_file = " & Scrub(HoldSelItem) & ", "
  strSql = strSql & "imported_table.update_date_time = Now(), " & _
                    "imported_table.update_user = " & Scrub(Environ("USERNAME")) & ", " & _
                    "imported_table.update_program = 'ImportFID', " & _
                    "imported_table.update_file = " & Scrub(HoldSelItem) & ", "
End If

If Len(strSql) > 0 Then strSql = Left(strSql, Len(strSql) - 2) & ";"

DebugPrint ("Step #2 'AddTableFields' - " & strSql)
DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True

Set fld = Nothing
Set tdf = Nothing
Set target = Nothing
Set db = Nothing
Set prop = Nothing

End Function

Private Function UserNameWindows() As String
Dim UserName  As String
     UserName = Environ("USERNAME")
     UserNameWindows = UserName
End Function




Private Function Create_Imported_Table_Errors_SQL()

Dim strSql  As String
Dim errMsg As String, errResponse As Long



On Error GoTo Error_Handler
DoCmd.SetWarnings False
strSql = "CREATE TABLE imported_table_errors (" _
       & "file_name                       TEXT(255), " _
       & "excel_work_sheet_name           TEXT(255), " _
       & "source_file_primary_key_heading TEXT(255), " _
       & "source_file_primary_key_value   TEXT(255), " _
       & "error_field_cell_number        TEXT(255), " _
       & "error_field_heading             TEXT(255), " _
       & "error_field_value               TEXT(255), " _
       & "error_message                   TEXT(255), " _
       & "file_name_full                  TEXT(255)) "
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True

aWorkSheetErrorFound = True       ' Set this flag to be used later.
aNoImportErrorsWereFound = False ' Set this flag to be used later to indicate an error occured.

Exit Function

Error_Handler:
  
  Select Case err.Number
  Case 3010
    Resume Next
  Case 3211
    errMsg = "Error number: " & Str(err.Number) & vbNewLine & _
             "Source: " & err.source & vbNewLine & _
             "Description: " & err.Description
    errResponse = MsgBox(errMsg, vbRetryCancel)
    DebugPrint (Chr(13) & "****" & errMsg & Chr(13))
    If errResponse = 4 Then Resume
    
    End
  Case Else
    errMsg = "Error number: " & Str(err.Number) & vbNewLine & _
             "Source: " & err.source & vbNewLine & _
             "Description: " & err.Description
    MsgBox (errMsg)
    Debug.Print (Chr(13) & "****" & errMsg & Chr(13))
    End
    Resume
  End Select
End Function
  

Private Sub Process_Excel_Data_Error_Check(ByVal fldPtr As Long, ByVal WorkSheetName As String)

Dim I As Long, J As Long
Dim rst As dao.Recordset
Dim strSql As String
Dim HOLD_KEY_VALUES     As String
Dim Entire_Worksheet_Should_be_Rejected  As Boolean
Entire_Worksheet_Should_be_Rejected = False

If aField_Name_Output(fldPtr) = "" Then Exit Sub   ' This is not an output field.
If aData_Type(fldPtr) = "" Then Exit Sub

 'Build the select SQL string
 strSql = "SELECT"
 
 'Build the SELECT for all rows with this error....
 Dim Hold_Key_Heading   As String
 Dim Hold_Key_Num       As Long
 Dim ErrorMsg   As String
 For I = 1 To numOfSpecs   ' Get the key field values....
    If aDup_Key_Field(I) Then
       Hold_Key_Num = Hold_Key_Num + 1
       '  Need to give error msgbox if more than 12
       If Hold_Key_Num > 12 Then
          ErrorMsg = "Too many key fields specified for a table.  Cannot be more than 12 fields specified for " & WorkSheetName
          MsgBox (ErrorMsg)
          Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
          End
       End If
       If Hold_Key_Num > 1 Then Hold_Key_Heading = Hold_Key_Heading & " / "
       Hold_Key_Heading = Hold_Key_Heading & aExcel_Heading_Text(I)
       If Hold_Key_Num > 1 Then strSql = strSql & ","
       strSql = strSql & " imported_table.F" & xlColNum(aExcel_Column_Number(I)) & " AS key" & Hold_Key_Num
       strSql = strSql & ", imported_table." & Br(aField_Name_Output(I)) & " AS akey" & Hold_Key_Num
    End If
 Next I
 strSql = strSql & ", imported_table.row_contains_error"
 strSql = strSql & ", imported_table.reject_err_row"
 
 strSql = strSql & ", imported_table.F" & xlColNum(aExcel_Column_Number(fldPtr)) & " AS bad_data"
 strSql = strSql & ", imported_table." & Br(aField_Name_Output(fldPtr)) & " AS abad_data"
 strSql = strSql & ", imported_table.[" & aField_Name_Output(fldPtr) & "_err] AS abad_data_err"
 
 strSql = strSql & ", imported_table.F" & xlColNum(aExcel_Column_Number(numOfSpecs)) & " AS row_number"
 
 strSql = strSql & " FROM imported_table WHERE ((Not (imported_table.F" & _
             xlColNum(aExcel_Column_Number(fldPtr)) & ")='blanks') AND ((imported_table." & _
             Br(aField_Name_Output(fldPtr)) & ") Is Null));"
 DebugPrint ("Process_Excel_Data_Error_Check=" & strSql)
 
 Set rst = Application.CurrentDb.OpenRecordset(strSql)  ' Open recordset with intent to edit.
 If rst.RecordCount = 0 Then
    rst.Close
    Exit Sub
 End If
 'DebugPrint (strSql)
 Create_Imported_Table_Errors_SQL  ' Create table to hold found errors.
 
'   Top of the READ loop for Import Table with field errors....
 Do
    HOLD_KEY_VALUES = ""
    If Hold_Key_Num > 0 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & rst!key1
    If Hold_Key_Num > 1 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key2
    If Hold_Key_Num > 2 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key3
    If Hold_Key_Num > 3 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key4
    If Hold_Key_Num > 4 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key5
    If Hold_Key_Num > 5 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key6
    If Hold_Key_Num > 6 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key7
    If Hold_Key_Num > 7 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key8
    If Hold_Key_Num > 8 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key9
    If Hold_Key_Num > 9 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key10
    If Hold_Key_Num > 10 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key11
    If Hold_Key_Num > 11 Then HOLD_KEY_VALUES = HOLD_KEY_VALUES & " / " & rst!key12
 
    Dim imported_table_errors As dao.Recordset
    Dim HoldErrMsg      As String
    Set imported_table_errors = CurrentDb.OpenRecordset("SELECT * FROM [imported_table_errors]")  ' Intend to edit
    imported_table_errors.AddNew
    imported_table_errors![file_name] = HoldFileName
    imported_table_errors![file_name_full] = HoldSelItem
    imported_table_errors![excel_work_sheet_name] = WorkSheetName
    imported_table_errors![source_file_primary_key_heading] = Hold_Key_Heading
    imported_table_errors![source_file_primary_key_value] = HOLD_KEY_VALUES
    imported_table_errors![error_field_cell_number] = aExcel_Column_Number(fldPtr) & rst!row_number
    imported_table_errors![error_field_heading] = aExcel_Heading_Text(fldPtr)
    imported_table_errors![error_field_value] = rst!bad_data
    
    HoldErrMsg = "Data is invalid " & aData_Type(fldPtr) & " format."
    If aReject_Err_Rows(fldPtr) Then HoldErrMsg = _
             "Entire Row is Rejected. Data is invalid " & aData_Type(fldPtr) & " format."
    If aReject_Err_File(fldPtr) Then HoldErrMsg = _
            "Entire Worksheet-" & HoldFileName & "/" & WorkSheetName & " is Rejected. Data is invalid " & aData_Type(fldPtr) & " format."
    imported_table_errors![error_message] = HoldErrMsg
    
    imported_table_errors.Update
    imported_table_errors.Close
    Set imported_table_errors = Nothing
    rst.Edit
    rst!row_contains_error = "Y"
    rst!abad_data_err = HoldErrMsg
    Dim HOLD_aReject_Err_Rows As Boolean
    HOLD_aReject_Err_Rows = aReject_Err_Rows(fldPtr)
    If aReject_Err_Rows(fldPtr) Then rst!reject_err_row = "Y"
   ' ***** rst!reject_err_row = "Y"  '  Always reject entire row if there is an error.
    If aReject_Err_File(fldPtr) Then Entire_Worksheet_Should_be_Rejected = True

    rst.Update
    rst.MoveNext

    If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
rst.Close
Set rst = Nothing

If Entire_Worksheet_Should_be_Rejected Then
  strSql = "UPDATE imported_table SET imported_table.reject_err_row = 'Y';"
  DoCmd.RunSQL (strSql) ' Mark all rows to be rejected due to error (When spec says entire worksheet should be rejected)
End If



End Sub

Private Sub Process_Field_Update_Data_Error_Check(rst As Recordset, _
                                          fldPtr As Long, _
                                          target As Variant, _
                                          source As Variant, _
                                          err As String, _
                                          WorkSheetName As String)

Dim I As Long, J As Long
Dim imported_table_errors As dao.Recordset
Dim strSql As String

If aField_Name_Output(fldPtr) = "" Then Exit Sub   ' This is not an output field.
If target = source Then Exit Sub                   ' The fields are equal, don't report as error.
If err <> "" Then Exit Sub                         ' If a previous error has been reported, then return.

Create_Imported_Table_Errors_SQL  ' Create table to hold found errors.
    
Set imported_table_errors = CurrentDb.OpenRecordset("SELECT * FROM [imported_table_errors]") ' Intend to edit.
imported_table_errors.AddNew
imported_table_errors![file_name] = HoldFileName
imported_table_errors![file_name_full] = HoldSelItem
imported_table_errors![excel_work_sheet_name] = WorkSheetName
imported_table_errors![source_file_primary_key_heading] = aExcel_Heading_Text(fldPtr)
imported_table_errors![source_file_primary_key_value] = rst!key_fld
imported_table_errors![error_field_cell_number] = aExcel_Column_Number(fldPtr) & rst!excel_row
imported_table_errors![error_field_heading] = aExcel_Heading_Text(fldPtr)
imported_table_errors![error_field_value] = source
    
imported_table_errors![error_message] = "Data VALUE in Excel Row has been rejected because of KEY or LOCK database Violations."
    
imported_table_errors.Update
imported_table_errors.Close
Set imported_table_errors = Nothing

strSql = "UPDATE imported_table SET imported_table.row_has_changed = 'N' WHERE ("
strSql = strSql & "imported_table.excel_row_number = " & rst!excel_row & ");"

DebugPrint ("Step 'row_has_changed=N' - " & strSql)
DoCmd.RunSQL (strSql)   '  Update matching rows...


End Sub

Private Function GetKeyValue(rst As Recordset, keyName As String)

If keyName = "key_1" Then GetKeyValue = rst!key_1
If keyName = "key_2" Then GetKeyValue = rst!key_2
If keyName = "key_3" Then GetKeyValue = rst!key_3
If keyName = "key_4" Then GetKeyValue = rst!key_4
If keyName = "key_5" Then GetKeyValue = rst!key_5
If keyName = "key_6" Then GetKeyValue = rst!key_6
If keyName = "key_7" Then GetKeyValue = rst!key_7
If keyName = "key_8" Then GetKeyValue = rst!key_8
If keyName = "key_9" Then GetKeyValue = rst!key_9
If keyName = "key_10" Then GetKeyValue = rst!key_10
If keyName = "key_11" Then GetKeyValue = rst!key_11
If keyName = "key_12" Then GetKeyValue = rst!key_12

End Function

Private Sub Setup_Row_Numbers_Definitions()
'  This routine will:
'   1)  Locate the Row_Number Column
'   2)  Add field definition to field specifications for Row Number...
'   3)  Delete ???xxxxx??? rows from the imported_table
'   4)  Add the row_count and excel_row_num fields. Then populate excel_row_num with the contents of the last column imported.

Dim I As Long, J As Long
Dim last_column   As String  ' Field that represents last column before row numbers.
Dim rowNumbersColumnName   As String
Dim holdRowFieldNumber   As Long
Dim strSql    As String

For I = 1 To UBound(aExcel_Column_Number)
  If Not CheckExists("F" & I) Then
    rowNumbersColumnName = "F" & (I - 1)
    holdRowFieldNumber = I - 1
    last_column = "F" & (I - 2)
    Exit For
  End If
Next I

DebugPrint ("rowNumbersColumnName=" & rowNumbersColumnName & "  " & "holdRowFieldNumber=" & holdRowFieldNumber & _
            "last_column=" & last_column)

If rowNumbersColumnName = "" Then
  MsgBox ("Error in Setup_Row_Numbers subroutine.  Row Number fields were not found.  Aborted import.")
  Debug.Print ("Error in Setup_Row_Numbers subroutine.  Row Number fields were not found.  Aborted import.")
  End
End If

'  Add field definition to field specifications for Row Number...
numOfSpecs = numOfSpecs + 1
J = numOfSpecs

aOutput_Table_Name(J) = ""
If J > 1 Then aOutput_Table_Name(J) = aOutput_Table_Name(J - 1)
aExcel_Column_Number(J) = xlColAlfa(holdRowFieldNumber)
aExcel_Heading_Text(J) = "Row_Number"
aField_Name_Output(J) = ""
aData_Type(J) = "Number"
aDup_Key_Field(J) = False
aReject_Err_File(J) = False
aReject_Err_Rows(J) = False
aAccept_Changes(J) = True
aAllowChange2Blank(J) = False
aSave_Date_Changed(J) = False

' Delete ???xxxxx??? rows from the imported_table
strSql = "DELETE imported_table.* FROM imported_table WHERE (((imported_table." & last_column & ")='???xxxxx???'));"
DebugPrint ("Step #3 'Setup_Row_Numbers' - " & strSql)
DoCmd.RunSQL (strSql)

'  Delete Empty Rows from the imported_table
Dim xx   As String
strSql = "DELETE imported_table.* FROM imported_table WHERE ("
For I = 1 To numOfSpecs - 1
  xx = "F" & xlColNum(aExcel_Column_Number(I))
  strSql = strSql & "((imported_table." & xx & ") Is Null Or (imported_table." & xx & ")='') AND "
Next I
strSql = Left(strSql, Len(strSql) - 5) & ");"
DebugPrint ("Step #3 'Setup_Row_Numbers' - " & strSql)
DoCmd.RunSQL (strSql)

Dim tdf       As TableDef
Dim db As dao.Database

Dim seqColumnNum     As String
seqColumnNum = "F" & xlColNum(aExcel_Column_Number(numOfSpecs))

' Add the row_count and excel_row_num fields. Then populate excel_row_num with the contents of the last column imported.
Set db = CurrentDb
Set tdf = db.TableDefs("imported_table")
On Error Resume Next
tdf.Fields.Append tdf.CreateField("excel_row_number", 7)     ' dbDouble
tdf.Fields.Append tdf.CreateField("row_count", 7)     ' dbDouble
On Error GoTo 0
Set tdf = Nothing
Set db = Nothing
strSql = "UPDATE " & "imported_table" & " SET " & "imported_table" & ".excel_row_number = [" & _
            "imported_table" & "].[" & seqColumnNum & "];"
DebugPrint ("Step #4 'Setup_Row_Numbers' - " & strSql)
DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True


End Sub

Private Sub Process_A_Worksheet(Worksheet_Name As String, Output_Table As String)

Dim I As Long, J As Long, k As Long
Dim strSql     As String
Dim db As dao.Database
Set db = CurrentDb
Dim rst As dao.Recordset
Dim fName   As String

Call GetImportSpecifications(Worksheet_Name, Output_Table)
If numOfSpecs = 0 Then Exit Sub

Call printImportSpecs(HoldFileName, Worksheet_Name)

If Not Specification_Is_Valid(Worksheet_Name, Output_Table) Then
  MsgBox ("Import for " & Worksheet_Name & "/" & Output_Table & " is terminated due to errors in the import specification.")
  Debug.Print ("Import for " & Worksheet_Name & "/" & Output_Table & " is terminated due to errors in the import specification.")
  End
End If

Dim prepedFile   As String
prepedFile = PrepXL(HoldSelItem, Worksheet_Name)

Debug.Print (" Process the Worksheet-" & Worksheet_Name)
DelTbl ("imported_table")

Dim holdRowNumberColumn       As Long
holdRowNumberColumn = Find_Row_Number_Column(prepedFile, Worksheet_Name)
If holdRowNumberColumn = 0 Then
  ErrorMsg = "Import for " & Worksheet_Name & "/" & Output_Table & " is terminated due to problem in the Prep_XL routine." & _
          Chr(13) & Chr(13) & "'Row_Number' column was not found to have been added to Excel spreadsheet." & _
          Chr(13) & Chr(13) & "Process will terminate"
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

'Step #1 - Transfer the data from the spreadsheet to imported table
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "imported_table", _
        prepedFile, False, Worksheet_Name & "!"
Kill (prepedFile)

'Step #1a - Delete any extra columns (after Row_Number) from the import_table......
For I = holdRowNumberColumn + 1 To 150
  fName = "F" & I
  If Not FieldExists(fName, "imported_table") Then Exit For
  Call Remove_Table_Field(fName, "imported_table")
Next I

'Step #1aa - Delete all NULL rows
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aData_Type(I) = "Date" Then
    strSql = "UPDATE imported_table SET imported_table.F" & xlColNum(aExcel_Column_Number(I)) & " = Null " _
           & "WHERE (((imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ")=""1/0/1900"")) OR (((imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ")=""12/31/1899""));"
    DebugPrint ("Step #1aa - " & strSql)
    DoCmd.RunSQL (strSql)   '  Change these null date values to actual NULL value...
  End If
Next I

'Step #1b - Delete all NULL rows
If holdRowNumberColumn = 1 Then GoTo skipStep1b
strSql = "DELETE imported_table.* FROM imported_table WHERE ("
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then _
    strSql = strSql & "((imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ") Is Null) AND "
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5) & ");"
DebugPrint ("Step #1b - " & strSql)
DoCmd.RunSQL (strSql)   '  Delete all NULL rows...
skipStep1b:


'Step #1c - Delete the Excel heading line from the import_table......
strSql = "DELETE imported_table.* FROM imported_table WHERE ("
For I = 1 To numOfSpecs
  If aExcel_Heading_Text(I) <> "" Then
    strSql = strSql & _
      "((imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ") = '" & aExcel_Heading_Text(I) & "') AND "
  End If
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5) & ");"
DebugPrint ("Step #1c - " & strSql)
DoCmd.RunSQL (strSql)

'Step #1d - Add an additional Specification Item to the numOfSpecs to capture the excel Row_Number column
Call Setup_Row_Numbers_Definitions

'Step #1e - Eliminate all dups from the imported_table (when keys are UNIQUE)
If aUnique_Dup_Key_Field Then _
  Call Merge_Duplicates("imported_table", Output_Table)

'Step #2 - Add SQL fields to imported_table as required, if Update Stamp is required, add update stamp fields to target table.
Call AddTableFields(Output_Table)

'Step #3 - Populate import_table named data fields and check for excel data errors "null's", generate all error messages
'         Mark "reject_err_row" if entire row should be deleted.
For I = 1 To numOfSpecs  ' Process all errors for defined fields.
   Call Process_Excel_Data_Error_Check(I, Worksheet_Name)  ' Process all excel data errors for this field....
Next I

'Step #3a - Examine the "imported_table" and fill all null fields with 0 for the numeric fields and blanks for string fields.
'           so that JOINs will not have to handle NULL values later.
Dim hDefaultNullValue     As String
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    hDefaultNullValue = """ """
    If aData_Type(I) = "Number" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Date" Then hDefaultNullValue = "1/1/1900"
    strSql = "UPDATE imported_table SET imported_table." & aField_Name_Output(I) & " = " & hDefaultNullValue & " WHERE (((imported_table." & aField_Name_Output(I) & ") Is Null));"
    DebugPrint ("Step #3a - " & strSql)
    DoCmd.RunSQL (strSql)   '  Now fill in all NULL values.
  End If
Next I


'Step #3b - Examine the "target" table and fill all null KEY fields with 0 for the numeric fields and blanks for string fields.
'           so that JOINs will not have to handle NULL values later.
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aDup_Key_Field(I) Then
    hDefaultNullValue = """ """
    If aData_Type(I) = "Number" Then hDefaultNullValue = "0"
    If aData_Type(I) = "Date" Then hDefaultNullValue = "0"
    strSql = "UPDATE " & Output_Table & " SET " & Output_Table & "." & aField_Name_Output(I) & " = " & _
      hDefaultNullValue & " WHERE (((" & Output_Table & "." & aField_Name_Output(I) & ") Is Null));"
    DebugPrint ("Step #3b - " & strSql)
    DoCmd.RunSQL (strSql)   '  Now fill in all NULL values.
  End If
Next I



'Step #4 - Delete any entire rows that have errors and that have been marked for error rejections.

strSql = "DELETE  imported_table.* FROM imported_table WHERE (((imported_table.reject_err_row)='Y'));"
DebugPrint ("Step #4 - " & strSql)
DoCmd.RunSQL (strSql)   '  Now delete all of the marked errors that qualify for delection from the input file.

' At this point, all appropriately marked rows have been deleted from the "imported_table" and
' we are ready now to add new rows and update existing rows in the target table.

'Step #5 - Mark all matching rows with matched_target_table = "Y" using an INNER JOIN to identify new rows that need
'          to be appended.
strSql = "UPDATE imported_table SET imported_table.matched_target_table = 'N';"
DoCmd.RunSQL (strSql)   '  Mark all rows...
strSql = "UPDATE imported_table INNER JOIN " & Output_Table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & _
       "(imported_table." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
  End If
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET imported_table.matched_target_table = 'Y', "
If aMark_active_flag = "Imported" Or aMark_active_flag = "Changed" Then
  strSql = strSql & Output_Table & "." & Br(aActive_flag_name) & " = Yes, "
End If
strSql = strSql & Output_Table & ".excel_row_number = [imported_table].[excel_row_number];"
DebugPrint ("Step #5 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all matching rows...

'Step #5a - When key values are NOT UNIQUE, all matching rows in the target table must first be deleted.
If Not aUnique_Dup_Key_Field Then _
  Call Delete_Matching_From_Target("imported_table", Output_Table)

'Step #6 - Set all matched_target_table = "N" error'ed field values to null to avoid bad data from being added to the table.
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = "UPDATE imported_table SET imported_table." & Br(aField_Name_Output(I)) & " = Null " & _
             "WHERE (((imported_table.[" & aField_Name_Output(I) & "_err]) Is Not Null And (imported_table.[" & aField_Name_Output(I) & "_err])<>''));"
    DebugPrint ("Step #6 - '" & Br(aField_Name_Output(I)) & "' " & strSql)
    DoCmd.RunSQL (strSql)   '  Clear data on errored fields...
  End If
Next I

'Step #7 - Append new records with matched_target_table = "N"
strSql = "INSERT INTO " & Output_Table & " (excel_row_number, "
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then strSql = strSql & Br(aField_Name_Output(I)) & ", "
Next I
If Add_Date_Changed_to_Rows Then _
  strSql = strSql & "update_date_time, update_user, update_program, update_file, create_date_time, create_user, create_program, create_file ) "
If Not Add_Date_Changed_to_Rows Then strSql = Left(strSql, Len(strSql) - 2) & " ) "

strSql = strSql & "SELECT imported_table.excel_row_number, "
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then _
   strSql = strSql & "imported_table." & Br(aField_Name_Output(I)) & ", "
Next I
If Add_Date_Changed_to_Rows Then
  strSql = strSql & "imported_table.update_date_time, imported_table.update_user, imported_table.update_program, imported_table.update_file, "
  strSql = strSql & "imported_table.create_date_time, imported_table.create_user, imported_table.create_program, imported_table.create_file, "
End If
strSql = Left(strSql, Len(strSql) - 2) & " "
strSql = strSql & "FROM imported_table WHERE (((imported_table.matched_target_table)='N'));"
DebugPrint ("Step #7 - " & strSql)
DoCmd.RunSQL (strSql)   '  Append rows...

'Step #7a - Mark all active flags = Yes using an INNER JOIN to identify new rows added that need to be marked active.
strSql = "UPDATE imported_table INNER JOIN " & Output_Table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & _
       "(imported_table." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
  End If
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET " & Output_Table & "." & Br(aActive_flag_name) & " = Yes "
strSql = strSql & "WHERE (((imported_table.matched_target_table)='N'));"
If aMark_active_flag = "New" Or aMark_active_flag = "Imported" Then
  DebugPrint ("Step #7a - " & strSql)
  DoCmd.RunSQL (strSql)   '  Mark Active Flags on all New rows...
End If

'Step 8 - evaluate the aAccept_Changes boolean to change the field to avoid an updated value being accepted in the data.
Dim Hold_Update_Clause     As String
Hold_Update_Clause = "UPDATE imported_table INNER JOIN " & Output_Table & " ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

'  Now go field by field to fill in imported_table with original values from the target table....
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And Not aAccept_Changes(I) Then
    strSql = Hold_Update_Clause & _
      "SET imported_table." & Br(aField_Name_Output(I)) & " = [" & Output_Table & "].[" & aField_Name_Output(I) & "] "
    If aData_Type(I) = "" Then strSql = strSql & _
         "WHERE (((" & Output_Table & "." & Br(aField_Name_Output(I)) & ") Is Not Null And (" & _
           Output_Table & "." & Br(aField_Name_Output(I)) & ")<>''));"
    If aData_Type(I) <> "" Then strSql = strSql & _
        "WHERE (" & Output_Table & "." & Br(aField_Name_Output(I)) & " Is Not Null);"
              
    DebugPrint ("Step #8 - " & strSql)
    DoCmd.RunSQL (strSql)   '  Pull the target fields and replace source fields...
  End If
Next I

'Step 8a - Find all "errored fields" and update the "imported_table" source field to match the target field....
Hold_Update_Clause = "UPDATE imported_table INNER JOIN " & Output_Table & " ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

'  Now go field by field to fill in imported_table with original values from the target table....
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = Hold_Update_Clause & _
      "SET imported_table." & Br(aField_Name_Output(I)) & " = [" & Output_Table & "].[" & aField_Name_Output(I) & "] "
    strSql = strSql & "WHERE (((imported_table.[" & aField_Name_Output(I) & "_err]) Is Not Null And (imported_table.[" & aField_Name_Output(I) & "_err])<>''));"
    DebugPrint ("Step #8a - " & strSql)
    DoCmd.RunSQL (strSql)   '  Pull the target fields and replace source fields...
  End If
Next I

'Step 8b - evaluate the aAllowChange2Blank boolean to avoid existing values changing to blank or zero when the field is empty in the excel spreadsheet.
'          If aAllowChange2Blank is false, then replace field values in the imported_table when field is null or blank.
Hold_Update_Clause = "UPDATE imported_table INNER JOIN " & Output_Table & " ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

'  Now go field by field to fill in imported_table with original values from the target table....
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And Not aAllowChange2Blank(I) Then
    strSql = Hold_Update_Clause & _
      "SET imported_table." & Br(aField_Name_Output(I)) & " = [" & Output_Table & "].[" & aField_Name_Output(I) & "] "
    
    strSql = strSql & _
        "WHERE (((imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ") Is Null Or (imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ")=''));"
              
    DebugPrint ("Step #8b - " & strSql)
    DoCmd.RunSQL (strSql)   '  Pull the target fields and replace source fields...
  End If
Next I



'Step 9 - Evaluate aSave_Date_Changed boolean and update the update stamp fields if the field value has changed. row_has_changed
Dim fieldValueCount   As Long
fieldValueCount = 0
strSql = "UPDATE imported_table INNER JOIN " & Output_Table & " ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

strSql = strSql & "SET imported_table.row_has_changed = 'Y' WHERE "
For I = 1 To numOfSpecs
  If aSave_Date_Changed(I) And aField_Name_Output(I) <> "" Then
    strSql = strSql & "([" & Output_Table & "].[" & aField_Name_Output(I) & "]<>[imported_table].[" & _
                                                    aField_Name_Output(I) & "]) OR "
    fieldValueCount = fieldValueCount + 1
  End If
Next I
If Right(strSql, 4) = " OR " Then strSql = Left(strSql, Len(strSql) - 4)
strSql = strSql & ";"
If Add_Date_Changed_to_Rows And fieldValueCount > 0 Then
  DebugPrint ("Step #9 - " & strSql)
  DoCmd.RunSQL (strSql)   '  Update CHANGED rows with Change Stamp
End If

'Step 10 - Actually update all source fields from the imported_table to the target table with the new field values.
fieldValueCount = 0
strSql = "UPDATE " & Output_Table & " INNER JOIN imported_table ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

strSql = strSql & "SET "
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = strSql & _
      Output_Table & "." & Br(aField_Name_Output(I)) & " = [imported_table].[" & aField_Name_Output(I) & "], "
    fieldValueCount = fieldValueCount + 1
  End If
Next I
If Right(strSql, 2) = ", " Then strSql = Left(strSql, Len(strSql) - 2)
strSql = strSql & ";"

If fieldValueCount > 0 Then
  DebugPrint ("Step #10 - " & strSql)
  'db.Execute (strSql)
  DoCmd.SetWarnings False
  DoCmd.RunSQL (strSql)   '  Update all rows...
  DoCmd.SetWarnings True
End If

'Step 11 - Verify that all fields were updated with new values.....  Use a LEFT JOIN to find ALL rows from imported_table
'          with all matching rows from the Target table.  Then all fields can be verified to ensure that the import worked.
strSql = "SELECT "

For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then _
    strSql = strSql & "[imported_table].[" & aField_Name_Output(I) & "] & ' / ' & "
Next I
If Right(strSql, 12) = "] & ' / ' & " Then strSql = Left(strSql, Len(strSql) - 11) & " AS key_fld, "

strSql = strSql & "[imported_table].[F" & xlColNum(aExcel_Column_Number(numOfSpecs)) & "] as excel_row, "

'  These key names will be used later to clear the row_has_changed field if there is a row level issue.
Dim holdKeyCtr   As Long
For I = LBound(holdKeyNames) To UBound(holdKeyNames)
  holdKeyNames(I) = ""  ' Initialize key names array
Next I
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And aDup_Key_Field(I) Then
    holdKeyCtr = holdKeyCtr + 1
    holdKeyNames(I) = "key_" & holdKeyCtr
    strSql = strSql & "imported_table." & Br(aField_Name_Output(I)) & " AS key_" & holdKeyCtr & ", "
  End If
Next I
 
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" Then
    strSql = strSql & Output_Table & "." & Br(aField_Name_Output(I)) & " AS tar_f" & xlColNum(aExcel_Column_Number(I)) & ", "
    strSql = strSql & "imported_table." & Br(aField_Name_Output(I)) & " AS src_f" & xlColNum(aExcel_Column_Number(I)) & ", "
    strSql = strSql & "imported_table.[" & aField_Name_Output(I) & "_err] AS err_f" & xlColNum(aExcel_Column_Number(I)) & ", "
  End If
Next I
strSql = Left(strSql, Len(strSql) - 2)  ' strip off the last comma.
strSql = strSql & " FROM imported_table LEFT JOIN " & Output_Table & " ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

DebugPrint ("Step #11 - " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
   rst.Close
   Set rst = Nothing
   Exit Sub
End If
'  Now loop thru the query results and look for examples of mismatched data....
Do
  For I = 1 To numOfSpecs
    If aField_Name_Output(I) <> "" Then _
      Call Report_Cells_Not_Updated_Errors(xlColNum(aExcel_Column_Number(I)), rst, Worksheet_Name, Output_Table)
  Next I

  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:

'Step 12 - Evaluate imported_data row_has_changed = 'Y' and update the update stamp fields if the field value has changed.
strSql = "UPDATE imported_table INNER JOIN " & Output_Table & " ON " & _
  "(imported_table.excel_row_number = " & Output_Table & ".excel_row_number) "

strSql = strSql & "SET " & _
                   Output_Table & ".update_date_time = [imported_table].[update_date_time], " & _
                   Output_Table & ".update_user = [imported_table].[update_user], " & _
                   Output_Table & ".update_program = [imported_table].[update_program], " & _
                   Output_Table & ".update_file = [imported_table].[update_file] " & _
            "WHERE imported_table.row_has_changed = 'Y';"
If Add_Date_Changed_to_Rows Then
  DebugPrint ("Step #12 - " & strSql)
  DoCmd.RunSQL (strSql)   '  Update CHANGED rows with Change Stamp
End If


Exit Sub

End Sub

Private Sub Report_Cells_Not_Updated_Errors(fldPtr As Long, rst As Recordset, Worksheet_Name As String, Output_Table As String)

Dim I As Long
I = fldPtr    '
' This routine will double check each field that may be different and call an update routine.

If I = 1 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f1, ""), Nz(rst!src_f1, ""), Nz(rst!err_f1, ""), Worksheet_Name)
If I = 2 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f2, ""), Nz(rst!src_f2, ""), Nz(rst!err_f2, ""), Worksheet_Name)
If I = 3 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f3, ""), Nz(rst!src_f3, ""), Nz(rst!err_f3, ""), Worksheet_Name)
If I = 4 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f4, ""), Nz(rst!src_f4, ""), Nz(rst!err_f4, ""), Worksheet_Name)
If I = 5 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f5, ""), Nz(rst!src_f5, ""), Nz(rst!err_f5, ""), Worksheet_Name)
If I = 6 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f6, ""), Nz(rst!src_f6, ""), Nz(rst!err_f6, ""), Worksheet_Name)
If I = 7 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f7, ""), Nz(rst!src_f7, ""), Nz(rst!err_f7, ""), Worksheet_Name)
If I = 8 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f8, ""), Nz(rst!src_f8, ""), Nz(rst!err_f8, ""), Worksheet_Name)
If I = 9 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f9, ""), Nz(rst!src_f9, ""), Nz(rst!err_f9, ""), Worksheet_Name)
If numOfSpecs <= 9 Then Exit Sub

If I = 10 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f10, ""), Nz(rst!src_f10, ""), Nz(rst!err_f10, ""), Worksheet_Name)
If I = 11 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f11, ""), Nz(rst!src_f11, ""), Nz(rst!err_f11, ""), Worksheet_Name)
If I = 12 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f12, ""), Nz(rst!src_f12, ""), Nz(rst!err_f12, ""), Worksheet_Name)
If I = 13 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f13, ""), Nz(rst!src_f13, ""), Nz(rst!err_f13, ""), Worksheet_Name)
If I = 14 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f14, ""), Nz(rst!src_f14, ""), Nz(rst!err_f14, ""), Worksheet_Name)
If I = 15 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f15, ""), Nz(rst!src_f15, ""), Nz(rst!err_f15, ""), Worksheet_Name)
If I = 16 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f16, ""), Nz(rst!src_f16, ""), Nz(rst!err_f16, ""), Worksheet_Name)
If I = 17 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f17, ""), Nz(rst!src_f17, ""), Nz(rst!err_f17, ""), Worksheet_Name)
If I = 18 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f18, ""), Nz(rst!src_f18, ""), Nz(rst!err_f18, ""), Worksheet_Name)
If I = 19 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f19, ""), Nz(rst!src_f19, ""), Nz(rst!err_f19, ""), Worksheet_Name)
If numOfSpecs <= 19 Then Exit Sub

If I = 20 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f20, ""), Nz(rst!src_f20, ""), Nz(rst!err_f20, ""), Worksheet_Name)
If I = 21 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f21, ""), Nz(rst!src_f21, ""), Nz(rst!err_f21, ""), Worksheet_Name)
If I = 22 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f22, ""), Nz(rst!src_f22, ""), Nz(rst!err_f22, ""), Worksheet_Name)
If I = 23 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f23, ""), Nz(rst!src_f23, ""), Nz(rst!err_f23, ""), Worksheet_Name)
If I = 24 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f24, ""), Nz(rst!src_f24, ""), Nz(rst!err_f24, ""), Worksheet_Name)
If I = 25 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f25, ""), Nz(rst!src_f25, ""), Nz(rst!err_f25, ""), Worksheet_Name)
If I = 26 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f26, ""), Nz(rst!src_f26, ""), Nz(rst!err_f26, ""), Worksheet_Name)
If I = 27 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f27, ""), Nz(rst!src_f27, ""), Nz(rst!err_f27, ""), Worksheet_Name)
If I = 28 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f28, ""), Nz(rst!src_f28, ""), Nz(rst!err_f28, ""), Worksheet_Name)
If I = 29 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f29, ""), Nz(rst!src_f29, ""), Nz(rst!err_f29, ""), Worksheet_Name)
If numOfSpecs <= 29 Then Exit Sub

If I = 30 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f30, ""), Nz(rst!src_f30, ""), Nz(rst!err_f30, ""), Worksheet_Name)
If I = 31 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f31, ""), Nz(rst!src_f31, ""), Nz(rst!err_f31, ""), Worksheet_Name)
If I = 32 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f32, ""), Nz(rst!src_f32, ""), Nz(rst!err_f32, ""), Worksheet_Name)
If I = 33 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f33, ""), Nz(rst!src_f33, ""), Nz(rst!err_f33, ""), Worksheet_Name)
If I = 34 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f34, ""), Nz(rst!src_f34, ""), Nz(rst!err_f34, ""), Worksheet_Name)
If I = 35 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f35, ""), Nz(rst!src_f35, ""), Nz(rst!err_f35, ""), Worksheet_Name)
If I = 36 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f36, ""), Nz(rst!src_f36, ""), Nz(rst!err_f36, ""), Worksheet_Name)
If I = 37 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f37, ""), Nz(rst!src_f37, ""), Nz(rst!err_f37, ""), Worksheet_Name)
If I = 38 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f38, ""), Nz(rst!src_f38, ""), Nz(rst!err_f38, ""), Worksheet_Name)
If I = 39 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f39, ""), Nz(rst!src_f39, ""), Nz(rst!err_f39, ""), Worksheet_Name)
If numOfSpecs <= 39 Then Exit Sub

If I = 40 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f40, ""), Nz(rst!src_f40, ""), Nz(rst!err_f40, ""), Worksheet_Name)
If I = 41 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f41, ""), Nz(rst!src_f41, ""), Nz(rst!err_f41, ""), Worksheet_Name)
If I = 42 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f42, ""), Nz(rst!src_f42, ""), Nz(rst!err_f42, ""), Worksheet_Name)
If I = 43 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f43, ""), Nz(rst!src_f43, ""), Nz(rst!err_f43, ""), Worksheet_Name)
If I = 44 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f44, ""), Nz(rst!src_f44, ""), Nz(rst!err_f44, ""), Worksheet_Name)
If I = 45 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f45, ""), Nz(rst!src_f45, ""), Nz(rst!err_f45, ""), Worksheet_Name)
If I = 46 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f46, ""), Nz(rst!src_f46, ""), Nz(rst!err_f46, ""), Worksheet_Name)
If I = 47 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f47, ""), Nz(rst!src_f47, ""), Nz(rst!err_f47, ""), Worksheet_Name)
If I = 48 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f48, ""), Nz(rst!src_f48, ""), Nz(rst!err_f48, ""), Worksheet_Name)
If I = 49 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f49, ""), Nz(rst!src_f49, ""), Nz(rst!err_f49, ""), Worksheet_Name)
If numOfSpecs <= 49 Then Exit Sub

If I = 50 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f50, ""), Nz(rst!src_f50, ""), Nz(rst!err_f50, ""), Worksheet_Name)
If I = 51 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f51, ""), Nz(rst!src_f51, ""), Nz(rst!err_f51, ""), Worksheet_Name)
If I = 52 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f52, ""), Nz(rst!src_f52, ""), Nz(rst!err_f52, ""), Worksheet_Name)
If I = 53 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f53, ""), Nz(rst!src_f53, ""), Nz(rst!err_f53, ""), Worksheet_Name)
If I = 54 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f54, ""), Nz(rst!src_f54, ""), Nz(rst!err_f54, ""), Worksheet_Name)
If I = 55 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f55, ""), Nz(rst!src_f55, ""), Nz(rst!err_f55, ""), Worksheet_Name)
If I = 56 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f56, ""), Nz(rst!src_f56, ""), Nz(rst!err_f56, ""), Worksheet_Name)
If I = 57 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f57, ""), Nz(rst!src_f57, ""), Nz(rst!err_f57, ""), Worksheet_Name)
If I = 58 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f58, ""), Nz(rst!src_f58, ""), Nz(rst!err_f58, ""), Worksheet_Name)
If I = 59 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f59, ""), Nz(rst!src_f59, ""), Nz(rst!err_f59, ""), Worksheet_Name)
If numOfSpecs <= 59 Then Exit Sub

If I = 60 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f60, ""), Nz(rst!src_f60, ""), Nz(rst!err_f60, ""), Worksheet_Name)
If I = 61 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f61, ""), Nz(rst!src_f61, ""), Nz(rst!err_f61, ""), Worksheet_Name)
If I = 62 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f62, ""), Nz(rst!src_f62, ""), Nz(rst!err_f62, ""), Worksheet_Name)
If I = 63 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f63, ""), Nz(rst!src_f63, ""), Nz(rst!err_f63, ""), Worksheet_Name)
If I = 64 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f64, ""), Nz(rst!src_f64, ""), Nz(rst!err_f64, ""), Worksheet_Name)
If I = 65 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f65, ""), Nz(rst!src_f65, ""), Nz(rst!err_f65, ""), Worksheet_Name)
If I = 66 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f66, ""), Nz(rst!src_f66, ""), Nz(rst!err_f66, ""), Worksheet_Name)
If I = 67 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f67, ""), Nz(rst!src_f67, ""), Nz(rst!err_f67, ""), Worksheet_Name)
If I = 68 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f68, ""), Nz(rst!src_f68, ""), Nz(rst!err_f68, ""), Worksheet_Name)
If I = 69 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f69, ""), Nz(rst!src_f69, ""), Nz(rst!err_f69, ""), Worksheet_Name)
If numOfSpecs <= 69 Then Exit Sub

If I = 70 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f70, ""), Nz(rst!src_f70, ""), Nz(rst!err_f70, ""), Worksheet_Name)
If I = 71 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f71, ""), Nz(rst!src_f71, ""), Nz(rst!err_f71, ""), Worksheet_Name)
If I = 72 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f72, ""), Nz(rst!src_f72, ""), Nz(rst!err_f72, ""), Worksheet_Name)
If I = 73 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f73, ""), Nz(rst!src_f73, ""), Nz(rst!err_f73, ""), Worksheet_Name)
If I = 74 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f74, ""), Nz(rst!src_f74, ""), Nz(rst!err_f74, ""), Worksheet_Name)
If I = 75 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f75, ""), Nz(rst!src_f75, ""), Nz(rst!err_f75, ""), Worksheet_Name)
If I = 76 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f76, ""), Nz(rst!src_f76, ""), Nz(rst!err_f76, ""), Worksheet_Name)
If I = 77 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f77, ""), Nz(rst!src_f77, ""), Nz(rst!err_f77, ""), Worksheet_Name)
If I = 78 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f78, ""), Nz(rst!src_f78, ""), Nz(rst!err_f78, ""), Worksheet_Name)
If I = 79 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f79, ""), Nz(rst!src_f79, ""), Nz(rst!err_f79, ""), Worksheet_Name)
If numOfSpecs <= 79 Then Exit Sub

If I = 80 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f80, ""), Nz(rst!src_f80, ""), Nz(rst!err_f80, ""), Worksheet_Name)
If I = 81 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f81, ""), Nz(rst!src_f81, ""), Nz(rst!err_f81, ""), Worksheet_Name)
If I = 82 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f82, ""), Nz(rst!src_f82, ""), Nz(rst!err_f82, ""), Worksheet_Name)
If I = 83 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f83, ""), Nz(rst!src_f83, ""), Nz(rst!err_f83, ""), Worksheet_Name)
If I = 84 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f84, ""), Nz(rst!src_f84, ""), Nz(rst!err_f84, ""), Worksheet_Name)
If I = 85 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f85, ""), Nz(rst!src_f85, ""), Nz(rst!err_f85, ""), Worksheet_Name)
If I = 86 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f86, ""), Nz(rst!src_f86, ""), Nz(rst!err_f86, ""), Worksheet_Name)
If I = 87 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f87, ""), Nz(rst!src_f87, ""), Nz(rst!err_f87, ""), Worksheet_Name)
If I = 88 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f88, ""), Nz(rst!src_f88, ""), Nz(rst!err_f88, ""), Worksheet_Name)
If I = 89 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f89, ""), Nz(rst!src_f89, ""), Nz(rst!err_f89, ""), Worksheet_Name)
If numOfSpecs <= 89 Then Exit Sub

If I = 90 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f90, ""), Nz(rst!src_f90, ""), Nz(rst!err_f90, ""), Worksheet_Name)
If I = 91 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f91, ""), Nz(rst!src_f91, ""), Nz(rst!err_f91, ""), Worksheet_Name)
If I = 92 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f92, ""), Nz(rst!src_f92, ""), Nz(rst!err_f92, ""), Worksheet_Name)
If I = 93 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f93, ""), Nz(rst!src_f93, ""), Nz(rst!err_f93, ""), Worksheet_Name)
If I = 94 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f94, ""), Nz(rst!src_f94, ""), Nz(rst!err_f94, ""), Worksheet_Name)
If I = 95 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f95, ""), Nz(rst!src_f95, ""), Nz(rst!err_f95, ""), Worksheet_Name)
If I = 96 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f96, ""), Nz(rst!src_f96, ""), Nz(rst!err_f96, ""), Worksheet_Name)
If I = 97 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f97, ""), Nz(rst!src_f97, ""), Nz(rst!err_f97, ""), Worksheet_Name)
If I = 98 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f98, ""), Nz(rst!src_f98, ""), Nz(rst!err_f98, ""), Worksheet_Name)
If I = 99 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f99, ""), Nz(rst!src_f99, ""), Nz(rst!err_f99, ""), Worksheet_Name)
If numOfSpecs <= 99 Then Exit Sub

If I = 100 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f100, ""), Nz(rst!src_f100, ""), Nz(rst!err_f100, ""), Worksheet_Name)
If I = 101 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f101, ""), Nz(rst!src_f101, ""), Nz(rst!err_f101, ""), Worksheet_Name)
If I = 102 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f102, ""), Nz(rst!src_f102, ""), Nz(rst!err_f102, ""), Worksheet_Name)
If I = 103 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f103, ""), Nz(rst!src_f103, ""), Nz(rst!err_f103, ""), Worksheet_Name)
If I = 104 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f104, ""), Nz(rst!src_f104, ""), Nz(rst!err_f104, ""), Worksheet_Name)
If I = 105 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f105, ""), Nz(rst!src_f105, ""), Nz(rst!err_f105, ""), Worksheet_Name)
If I = 106 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f106, ""), Nz(rst!src_f106, ""), Nz(rst!err_f106, ""), Worksheet_Name)
If I = 107 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f107, ""), Nz(rst!src_f107, ""), Nz(rst!err_f107, ""), Worksheet_Name)
If I = 108 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f108, ""), Nz(rst!src_f108, ""), Nz(rst!err_f108, ""), Worksheet_Name)
If I = 109 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f109, ""), Nz(rst!src_f109, ""), Nz(rst!err_f109, ""), Worksheet_Name)
If numOfSpecs <= 109 Then Exit Sub

If I = 110 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f110, ""), Nz(rst!src_f110, ""), Nz(rst!err_f110, ""), Worksheet_Name)
If I = 111 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f111, ""), Nz(rst!src_f111, ""), Nz(rst!err_f111, ""), Worksheet_Name)
If I = 112 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f112, ""), Nz(rst!src_f112, ""), Nz(rst!err_f112, ""), Worksheet_Name)
If I = 113 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f113, ""), Nz(rst!src_f113, ""), Nz(rst!err_f113, ""), Worksheet_Name)
If I = 114 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f114, ""), Nz(rst!src_f114, ""), Nz(rst!err_f114, ""), Worksheet_Name)
If I = 115 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f115, ""), Nz(rst!src_f115, ""), Nz(rst!err_f115, ""), Worksheet_Name)
If I = 116 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f116, ""), Nz(rst!src_f116, ""), Nz(rst!err_f116, ""), Worksheet_Name)
If I = 117 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f117, ""), Nz(rst!src_f117, ""), Nz(rst!err_f117, ""), Worksheet_Name)
If I = 118 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f118, ""), Nz(rst!src_f118, ""), Nz(rst!err_f118, ""), Worksheet_Name)
If I = 119 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f119, ""), Nz(rst!src_f119, ""), Nz(rst!err_f119, ""), Worksheet_Name)
If numOfSpecs <= 119 Then Exit Sub

If I = 120 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f120, ""), Nz(rst!src_f120, ""), Nz(rst!err_f120, ""), Worksheet_Name)
If I = 121 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f121, ""), Nz(rst!src_f121, ""), Nz(rst!err_f121, ""), Worksheet_Name)
If I = 122 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f122, ""), Nz(rst!src_f122, ""), Nz(rst!err_f122, ""), Worksheet_Name)
If I = 123 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f123, ""), Nz(rst!src_f123, ""), Nz(rst!err_f123, ""), Worksheet_Name)
If I = 124 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f124, ""), Nz(rst!src_f124, ""), Nz(rst!err_f124, ""), Worksheet_Name)
If I = 125 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f125, ""), Nz(rst!src_f125, ""), Nz(rst!err_f125, ""), Worksheet_Name)
If I = 126 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f126, ""), Nz(rst!src_f126, ""), Nz(rst!err_f126, ""), Worksheet_Name)
If I = 127 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f127, ""), Nz(rst!src_f127, ""), Nz(rst!err_f127, ""), Worksheet_Name)
If I = 128 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f128, ""), Nz(rst!src_f128, ""), Nz(rst!err_f128, ""), Worksheet_Name)
If I = 129 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f129, ""), Nz(rst!src_f129, ""), Nz(rst!err_f129, ""), Worksheet_Name)
If numOfSpecs <= 129 Then Exit Sub

If I = 130 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f130, ""), Nz(rst!src_f130, ""), Nz(rst!err_f130, ""), Worksheet_Name)
If I = 131 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f131, ""), Nz(rst!src_f131, ""), Nz(rst!err_f131, ""), Worksheet_Name)
If I = 132 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f132, ""), Nz(rst!src_f132, ""), Nz(rst!err_f132, ""), Worksheet_Name)
If I = 133 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f133, ""), Nz(rst!src_f133, ""), Nz(rst!err_f133, ""), Worksheet_Name)
If I = 134 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f134, ""), Nz(rst!src_f134, ""), Nz(rst!err_f134, ""), Worksheet_Name)
If I = 135 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f135, ""), Nz(rst!src_f135, ""), Nz(rst!err_f135, ""), Worksheet_Name)
If I = 136 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f136, ""), Nz(rst!src_f136, ""), Nz(rst!err_f136, ""), Worksheet_Name)
If I = 137 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f137, ""), Nz(rst!src_f137, ""), Nz(rst!err_f137, ""), Worksheet_Name)
If I = 138 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f138, ""), Nz(rst!src_f138, ""), Nz(rst!err_f138, ""), Worksheet_Name)
If I = 139 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f139, ""), Nz(rst!src_f139, ""), Nz(rst!err_f139, ""), Worksheet_Name)
If numOfSpecs <= 139 Then Exit Sub

If I = 140 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f140, ""), Nz(rst!src_f140, ""), Nz(rst!err_f140, ""), Worksheet_Name)
If I = 141 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f141, ""), Nz(rst!src_f141, ""), Nz(rst!err_f141, ""), Worksheet_Name)
If I = 142 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f142, ""), Nz(rst!src_f142, ""), Nz(rst!err_f142, ""), Worksheet_Name)
If I = 143 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f143, ""), Nz(rst!src_f143, ""), Nz(rst!err_f143, ""), Worksheet_Name)
If I = 144 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f144, ""), Nz(rst!src_f144, ""), Nz(rst!err_f144, ""), Worksheet_Name)
If I = 145 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f145, ""), Nz(rst!src_f145, ""), Nz(rst!err_f145, ""), Worksheet_Name)
If I = 146 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f146, ""), Nz(rst!src_f146, ""), Nz(rst!err_f146, ""), Worksheet_Name)
If I = 147 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f147, ""), Nz(rst!src_f147, ""), Nz(rst!err_f147, ""), Worksheet_Name)
If I = 148 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f148, ""), Nz(rst!src_f148, ""), Nz(rst!err_f148, ""), Worksheet_Name)
If I = 149 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f149, ""), Nz(rst!src_f149, ""), Nz(rst!err_f149, ""), Worksheet_Name)
If I = 150 Then Call Process_Field_Update_Data_Error_Check(rst, I, Nz(rst!tar_f150, ""), Nz(rst!src_f150, ""), Nz(rst!err_f150, ""), Worksheet_Name)

End Sub

'Dim strSql         As String
'Dim HoldSelItem    As String    ' Combined path and file name.
'Dim HoldFileName   As String    ' Only the file name.  Path is not included.
'Dim HoldFilePath   As String    ' Only the path name.

'  These Following array values are populated by GetImportSpecifications and includes only Active Items
'Dim numOfSpecs   As Long
'Dim Add_Date_Changed_to_Rows  As Boolean
'Dim aOutput_Table_Name(1 To 150) As String
'Dim aExcel_Column_Number(1 To 150) As String
'Dim aExcel_Heading_Text(1 To 150) As String
'Dim aField_Name_Output(1 To 150) As String
'Dim afield_err(1 To 150) As String
'Dim aData_Type(1 To 150) As String
'Dim aDup_Key_Field(1 To 150) As Boolean
'Dim aReject_Err_File(1 To 150) As Boolean
'Dim aReject_Err_Rows(1 To 150) As Boolean
'Dim aAccept_Changes(1 To 150) As Boolean
'Dim aSave_Date_Changed(1 To 150) As Boolean

'Dim aWorkSheetNames_len As Long ' For an xlsx file, this array will hold list of tabs.
'Dim aWorkSheetNames(1 To 150)  As String
'Dim aTablesFound(1 To 150)     As String


'INSERT INTO bank_account ( bank_account_id, bank, type_of_account, current_interest_rate, account_designator, veteran, account_number, depositor_account_title, fiduciary, last_reported_balance, last_balance_as_of_date )
'SELECT imported_table.bank_account_id, imported_table.bank, imported_table.type_of_account, imported_table.current_interest_rate, imported_table.account_designator, imported_table.veteran, imported_table.account_number, imported_table.depositor_account_title, imported_table.fiduciary, imported_table.last_reported_balance, imported_table.last_balance_as_of_date
'FROM imported_table
'WHERE (((imported_table.matched_target_table)='N'));


Private Sub Merge_Duplicates(imported_table As String, Output_Table As String)
'  This subroutine will merge duplicates in the "imported_table" based on the aDup_Key_Field from the Spec-ifications

Dim rst As dao.Recordset
Dim I As Long


'  Step 1 - Create a table to count the number occurances of the key values and then merge it into the imported_table
'           to populate the "row_count" field.
DelTbl ("imported_table_count")
strSql = "SELECT "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & "imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ", "
  End If
Next I
If strSql = "SELECT " Then
  ErrorMsg = Output_Table & " import specification has no Key Fields marked.  Correct this and restart the import."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  End
End If

strSql = Left(strSql, Len(strSql) - 2)
strSql = strSql & ", Count(imported_table.[F1]) AS countx INTO imported_table_count  FROM imported_table GROUP BY "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & "imported_table.F" & xlColNum(aExcel_Column_Number(I)) & ", "
  End If
Next I
strSql = Left(strSql, Len(strSql) - 2) & ";"
DebugPrint ("Step #1 - 'Merge_Duplicates' " & strSql)
DoCmd.RunSQL (strSql)

' Step 2 - Update the record counts in the imported_table
strSql = "UPDATE imported_table INNER JOIN imported_table_count ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then
    strSql = strSql & "(imported_table.F" & xlColNum(aExcel_Column_Number(I)) & _
        " = imported_table_count.F" & xlColNum(aExcel_Column_Number(I)) & ") AND "
  End If
Next I
strSql = Left(strSql, Len(strSql) - 4) & "SET imported_table.row_count = [imported_table_count].[countx];"
DebugPrint ("Step #2 - 'Merge_Duplicates' " & strSql)
DoCmd.RunSQL (strSql)

'Step 3 - Collect all of the records that have multiple keys that need to be merged.
strSql = "SELECT imported_table.* FROM imported_table WHERE (((imported_table.row_count) > 1)) " _
       & "ORDER BY imported_table.excel_row_number;"
       
Dim holdKeySql As String
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then holdKeySql = holdKeySql & "[F" & xlColNum(aExcel_Column_Number(I)) & "] & ',' & "
Next I
If Right(holdKeySql, 9) = " & ',' & " Then holdKeySql = Left(holdKeySql, Len(holdKeySql) - 9)

strSql = "SELECT imported_table.*, [F1] & ',' & [F2] AS tab_key FROM imported_table WHERE (((imported_table.row_count) > 1)) ORDER BY [F1] & ',' & [F2], imported_table.excel_row_number;"
DebugPrint ("Step #3 - 'Merge_Duplicates' " & strSql)
       
strSql = "SELECT imported_table.*, " & holdKeySql _
       & " AS tab_key FROM imported_table WHERE (((imported_table.row_count) > 1)) ORDER BY " _
       & holdKeySql & ", imported_table.excel_row_number;"
DebugPrint ("Step #4 - 'Merge_Duplicates' " & strSql)
       
Dim holdData(1 To 150) As String  '  This array will hold the contents of one row...
Dim lastKey As String, holdCount As Long, excelRowNumber As String
Dim LastRowNumber   As Double

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
   rst.Close
   Set rst = Nothing
   Exit Sub
End If
'  Now loop thru the query results and populate the array table
lastKey = "?"  ' Force Initial Break
Do
  If lastKey <> rst!tab_key Then Call Merge_Break(rst!tab_key, lastKey, holdCount, excelRowNumber, LastRowNumber)
    
  'Process an individual record.
  LastRowNumber = Nz(rst!excel_row_number, 0)
  Call merge_values(rst, holdCount, excelRowNumber)
  
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop
Loop
Finished_Do_Loop:
Call Merge_Break("?", lastKey, holdCount, excelRowNumber, LastRowNumber)  ' Take the final break....

' Delete all of the original rows that have been merged.......
strSql = "DELETE imported_table.row_count FROM imported_table WHERE (((imported_table.row_count)>1));"
DebugPrint ("'Merge_Break' " & strSql)
DoCmd.RunSQL (strSql)


rst.Close
Set rst = Nothing

Exit Sub
DelTbl ("imported_table_count")


End Sub
Private Sub Merge_Break(thisKey As String, _
                lastKey As String, _
                holdCount As Long, _
                excelRowNumber As String, _
                LastRowNumber As Double)
Dim I As Long, x As String

If lastKey = "?" Then GoTo First_Time
' Here we write out new merged row...

strSql = "INSERT INTO imported_table ( "
For I = 1 To (numOfSpecs)
  If holdMergedData(I) <> "#NULL#" Then strSql = strSql & "F" & I & ", "
Next I
strSql = strSql & "row_count, excel_row_number) VALUES ( "
For I = 1 To (numOfSpecs - 1)
  If holdMergedData(I) <> "#NULL#" Then strSql = strSql & Scrub(holdMergedData(I)) & ", "
Next I
strSql = strSql & "'" & excelRowNumber & "', " & holdCount & ", " & LastRowNumber & ");"
DebugPrint ("'Merge_Break' " & strSql)
DoCmd.RunSQL (strSql)

First_Time:
If thisKey = "?" Then Exit Sub

' Now prepare for the next set of records.
holdCount = 0
excelRowNumber = ""  ' Clear Previous Values
lastKey = thisKey
For I = 1 To numOfSpecs
  holdMergedData(I) = ""  ' Clear old values....
Next I

End Sub

Private Sub merge_values(rst As Recordset, holdCount As Long, excelRowNumber As String)

Dim holdData(1 To 150) As String
Dim I As Long

If excelRowNumber = "" Then
  excelRowNumber = Nz(rst!excel_row_number, "")
 Else
  excelRowNumber = excelRowNumber & ", " & Nz(rst!excel_row_number, "")
End If

holdCount = rst!row_count * -1

On Error GoTo Continue_Merge
holdData(1) = Nz(rst!F1, "#NULL#")
holdData(2) = Nz(rst!F2, "#NULL#")
holdData(3) = Nz(rst!f3, "#NULL#")
holdData(4) = Nz(rst!f4, "#NULL#")
holdData(5) = Nz(rst!f5, "#NULL#")
holdData(6) = Nz(rst!f6, "#NULL#")
holdData(7) = Nz(rst!f7, "#NULL#")
holdData(8) = Nz(rst!f8, "#NULL#")
holdData(9) = Nz(rst!f9, "#NULL#")

holdData(10) = Nz(rst!f10, "#NULL#")
holdData(11) = Nz(rst!f11, "#NULL#")
holdData(12) = Nz(rst!f12, "#NULL#")
holdData(13) = Nz(rst!f13, "#NULL#")
holdData(14) = Nz(rst!f14, "#NULL#")
holdData(15) = Nz(rst!f15, "#NULL#")
holdData(16) = Nz(rst!f16, "#NULL#")
holdData(17) = Nz(rst!f17, "#NULL#")
holdData(18) = Nz(rst!f18, "#NULL#")
holdData(19) = Nz(rst!f19, "#NULL#")

holdData(20) = Nz(rst!f20, "#NULL#")
holdData(21) = Nz(rst!f21, "#NULL#")
holdData(22) = Nz(rst!f22, "#NULL#")
holdData(23) = Nz(rst!f23, "#NULL#")
holdData(24) = Nz(rst!f24, "#NULL#")
holdData(25) = Nz(rst!f25, "#NULL#")
holdData(26) = Nz(rst!f26, "#NULL#")
holdData(27) = Nz(rst!f27, "#NULL#")
holdData(28) = Nz(rst!f28, "#NULL#")
holdData(29) = Nz(rst!f29, "#NULL#")

holdData(30) = Nz(rst!f30, "#NULL#")
holdData(31) = Nz(rst!f31, "#NULL#")
holdData(32) = Nz(rst!f32, "#NULL#")
holdData(33) = Nz(rst!f33, "#NULL#")
holdData(34) = Nz(rst!f34, "#NULL#")
holdData(35) = Nz(rst!f35, "#NULL#")
holdData(36) = Nz(rst!f36, "#NULL#")
holdData(37) = Nz(rst!f37, "#NULL#")
holdData(38) = Nz(rst!f38, "#NULL#")
holdData(39) = Nz(rst!f39, "#NULL#")

holdData(40) = Nz(rst!f40, "#NULL#")
holdData(41) = Nz(rst!f41, "#NULL#")
holdData(42) = Nz(rst!f42, "#NULL#")
holdData(43) = Nz(rst!f43, "#NULL#")
holdData(44) = Nz(rst!f44, "#NULL#")
holdData(45) = Nz(rst!f45, "#NULL#")
holdData(46) = Nz(rst!f46, "#NULL#")
holdData(47) = Nz(rst!f47, "#NULL#")
holdData(48) = Nz(rst!f48, "#NULL#")
holdData(49) = Nz(rst!f49, "#NULL#")

holdData(50) = Nz(rst!F50, "#NULL#")
holdData(51) = Nz(rst!F51, "#NULL#")
holdData(52) = Nz(rst!F52, "#NULL#")
holdData(53) = Nz(rst!F53, "#NULL#")
holdData(54) = Nz(rst!F54, "#NULL#")
holdData(55) = Nz(rst!F55, "#NULL#")
holdData(56) = Nz(rst!F56, "#NULL#")
holdData(57) = Nz(rst!F57, "#NULL#")
holdData(58) = Nz(rst!F58, "#NULL#")
holdData(59) = Nz(rst!F59, "#NULL#")

holdData(60) = Nz(rst!F60, "#NULL#")
holdData(61) = Nz(rst!F61, "#NULL#")
holdData(62) = Nz(rst!F62, "#NULL#")
holdData(63) = Nz(rst!F63, "#NULL#")
holdData(64) = Nz(rst!F64, "#NULL#")
holdData(65) = Nz(rst!F65, "#NULL#")
holdData(66) = Nz(rst!F66, "#NULL#")
holdData(67) = Nz(rst!F67, "#NULL#")
holdData(68) = Nz(rst!F68, "#NULL#")
holdData(69) = Nz(rst!F69, "#NULL#")

holdData(70) = Nz(rst!F70, "#NULL#")
holdData(71) = Nz(rst!F71, "#NULL#")
holdData(72) = Nz(rst!F72, "#NULL#")
holdData(73) = Nz(rst!F73, "#NULL#")
holdData(74) = Nz(rst!F74, "#NULL#")
holdData(75) = Nz(rst!F75, "#NULL#")
holdData(76) = Nz(rst!F76, "#NULL#")
holdData(77) = Nz(rst!F77, "#NULL#")
holdData(78) = Nz(rst!F78, "#NULL#")
holdData(79) = Nz(rst!F79, "#NULL#")

holdData(80) = Nz(rst!F80, "#NULL#")
holdData(81) = Nz(rst!F81, "#NULL#")
holdData(82) = Nz(rst!F82, "#NULL#")
holdData(83) = Nz(rst!F83, "#NULL#")
holdData(84) = Nz(rst!F84, "#NULL#")
holdData(85) = Nz(rst!F85, "#NULL#")
holdData(86) = Nz(rst!F86, "#NULL#")
holdData(87) = Nz(rst!F87, "#NULL#")
holdData(88) = Nz(rst!F88, "#NULL#")
holdData(89) = Nz(rst!F89, "#NULL#")

holdData(90) = Nz(rst!F90, "#NULL#")
holdData(91) = Nz(rst!F91, "#NULL#")
holdData(92) = Nz(rst!F92, "#NULL#")
holdData(93) = Nz(rst!F93, "#NULL#")
holdData(94) = Nz(rst!F94, "#NULL#")
holdData(95) = Nz(rst!F95, "#NULL#")
holdData(96) = Nz(rst!F96, "#NULL#")
holdData(97) = Nz(rst!F97, "#NULL#")
holdData(98) = Nz(rst!F98, "#NULL#")
holdData(99) = Nz(rst!F99, "#NULL#")

holdData(100) = Nz(rst!F100, "#NULL#")
holdData(101) = Nz(rst!F101, "#NULL#")
holdData(102) = Nz(rst!F102, "#NULL#")
holdData(103) = Nz(rst!F103, "#NULL#")
holdData(104) = Nz(rst!F104, "#NULL#")
holdData(105) = Nz(rst!F105, "#NULL#")
holdData(106) = Nz(rst!F106, "#NULL#")
holdData(107) = Nz(rst!F107, "#NULL#")
holdData(108) = Nz(rst!F108, "#NULL#")
holdData(109) = Nz(rst!F109, "#NULL#")

holdData(110) = Nz(rst!F110, "#NULL#")
holdData(111) = Nz(rst!F111, "#NULL#")
holdData(112) = Nz(rst!F112, "#NULL#")
holdData(113) = Nz(rst!F113, "#NULL#")
holdData(114) = Nz(rst!F114, "#NULL#")
holdData(115) = Nz(rst!F115, "#NULL#")
holdData(116) = Nz(rst!F116, "#NULL#")
holdData(117) = Nz(rst!F117, "#NULL#")
holdData(118) = Nz(rst!F118, "#NULL#")
holdData(119) = Nz(rst!F119, "#NULL#")

holdData(120) = Nz(rst!F120, "#NULL#")
holdData(121) = Nz(rst!F121, "#NULL#")
holdData(122) = Nz(rst!F122, "#NULL#")
holdData(123) = Nz(rst!F123, "#NULL#")
holdData(124) = Nz(rst!F124, "#NULL#")
holdData(125) = Nz(rst!F125, "#NULL#")
holdData(126) = Nz(rst!F126, "#NULL#")
holdData(127) = Nz(rst!F127, "#NULL#")
holdData(128) = Nz(rst!F128, "#NULL#")
holdData(129) = Nz(rst!F129, "#NULL#")

holdData(130) = Nz(rst!F130, "#NULL#")
holdData(131) = Nz(rst!F131, "#NULL#")
holdData(132) = Nz(rst!F132, "#NULL#")
holdData(133) = Nz(rst!F133, "#NULL#")
holdData(134) = Nz(rst!F134, "#NULL#")
holdData(135) = Nz(rst!F135, "#NULL#")
holdData(136) = Nz(rst!F136, "#NULL#")
holdData(137) = Nz(rst!F137, "#NULL#")
holdData(138) = Nz(rst!F138, "#NULL#")
holdData(139) = Nz(rst!F139, "#NULL#")

holdData(140) = Nz(rst!F140, "#NULL#")
holdData(141) = Nz(rst!F141, "#NULL#")
holdData(142) = Nz(rst!F142, "#NULL#")
holdData(143) = Nz(rst!F143, "#NULL#")
holdData(144) = Nz(rst!F144, "#NULL#")
holdData(145) = Nz(rst!F145, "#NULL#")
holdData(146) = Nz(rst!F146, "#NULL#")
holdData(147) = Nz(rst!F147, "#NULL#")
holdData(148) = Nz(rst!F148, "#NULL#")
holdData(149) = Nz(rst!F149, "#NULL#")
holdData(150) = Nz(rst!F150, "#NULL#")

Continue_Merge:
  On Error GoTo 0
  
For I = 1 To numOfSpecs  '  This loop will merge the values....
  If excelRowNumber = rst!excel_row_number Then _
     holdMergedData(I) = holdData(I)  ' Load the initial values.
          
  If Not aAllowChange2Blank(I) And (holdData(I) = "#NULL#" Or holdData(I) = "") Then _
    holdData(I) = holdMergedData(I) ' Don't allow blanks to survive the merge (except in the case of the first item.
     
  If aAccept_Changes(I) Then _
    holdMergedData(I) = holdData(I)  ' Load the changed values.
Next I
  
End Sub



Private Function Field_Type(aTable As String, aFieldName As String)
Dim objRecordset As ADODB.Recordset
Dim I As Long

Set objRecordset = New ADODB.Recordset
objRecordset.ActiveConnection = CurrentProject.Connection
objRecordset.Open ("SELECT " & aTable & ".* FROM " & aTable & ";")

For I = 0 To objRecordset.Fields.Count - 1
  If aFieldName = objRecordset.Fields.Item(I).Name Then GoTo Found_Field
Next I
Field_Type = aFieldName & " was not found in Table-" & aTable
Exit Function

Found_Field:
  If (objRecordset.Fields.Item(I).Type = 3) Or _
     (objRecordset.Fields.Item(I).Type = 20) Or _
     (objRecordset.Fields.Item(I).Type = 5) Or _
     (objRecordset.Fields.Item(I).Type = 6) Then
    Field_Type = "number"
    Exit Function
  End If

  If (objRecordset.Fields.Item(I).Type = 7) Then
    Field_Type = "date"
    Exit Function
  End If

  If (objRecordset.Fields.Item(I).Type = 11) Then
    Field_Type = "Yes/No"
    Exit Function
  End If

  If (objRecordset.Fields.Item(I).Type = 202) Or _
     (objRecordset.Fields.Item(I).Type = 203) Then
    Field_Type = ""
    Exit Function
  End If
  
  DebugPrint ("Unknown Field Type-" & objRecordset.Fields.Item(I).Type & "  for Table-" & aTable & "  Field-" & aFieldName)
  Field_Type = "Unknown Field Type-" & objRecordset.Fields.Item(I).Type

End Function


Private Function PrepXL(xlfile As String, wkSheet As String)

   Dim ImportPrep As String
   'Dim xlApp As Excel.Application
   'Dim xlBk As Excel.Workbook
   'Dim xlSht As Excel.Worksheet
   Dim xlApp As Object
   Dim xlBk As Object
   Dim xlSht As Object
   Dim ChkXL As String
   Dim I   As Long
      
   Dim holdXtension        As String    '  Get the file extension from the file path.
   I = Len(xlfile) - InStrRev(xlfile, ".")
   holdXtension = Right(xlfile, I)
   
   Dim HoldFilePath        As String
   I = InStrRev(xlfile, "\")
   HoldFilePath = Left(xlfile, I)
   
'   ImportPrep = CurrentProject.Path & "\ImportPrep." & holdXtension
   ImportPrep = HoldFilePath & "ImportPrep.xlsx"
   If Dir(ImportPrep) <> "" Then
      Kill (ImportPrep)
   End If
  ' FileCopy xlfile, ImportPrep
   DebugPrint ("xlfile=" & xlfile)
   DebugPrint ("ImportPrep=" & ImportPrep)
   Set xlApp = CreateObject("Excel.Application")
   Set xlBk = xlApp.Workbooks.Open(xlfile)
   Set xlSht = xlBk.Sheets(1)
   xlSht.Activate
   
   Dim lastCell    As String
   If holdXtension = "csv" Then
      xlApp.Sheets(1).Select
      xlApp.Selection.AutoFilter
     Else
      xlApp.Sheets(wkSheet).Select
      xlApp.Selection.AutoFilter     '  Turn off filters.
   End If
   With xlApp
   .ActiveCell.SpecialCells(xlLastCell).Select
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = " "
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .Range("A1").Select
   .Selection.End(xlToRight).Select
   .ActiveCell.Offset(0, 1).Range("A1").Select
   .ActiveCell.FormulaR1C1 = "Row_Number"
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = "2"
   .ActiveCell.Offset(1, 0).Range("A1").Select
   .ActiveCell.FormulaR1C1 = "=R[-1]C+1"
   .ActiveCell.Select
   .Selection.Copy
   .Range(.Selection, .ActiveCell.SpecialCells(xlLastCell)).Select
   .ActiveSheet.Paste
   .Application.CutCopyMode = False
   .Range("A1").Select
  End With
   
'    xlBk.Save
    xlBk.SaveAs FileName:=ImportPrep _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    xlBk.Close
    Set xlBk = Nothing
    Set xlApp = Nothing
    PrepXL = ImportPrep  ' Return back with the Prep'd file
End Function




Private Function CheckExists(ByVal strField As String) As Boolean

Dim objRecordset As ADODB.Recordset
Dim I As Long

Set objRecordset = New ADODB.Recordset
objRecordset.ActiveConnection = CurrentProject.Connection
objRecordset.Open ("imported_table")

'loop through table fields and check for a match
For I = 0 To objRecordset.Fields.Count - 1
  If strField = objRecordset.Fields.Item(I).Name Then  'exist function and return true
    CheckExists = True
    Exit Function
  End If
Next I

CheckExists = False  'return false
End Function

Private Sub Delete_Matching_From_Target(imported_table As String, Output_Table As String)
'  This subroutine will delete all of the Matching rows from the Target table to prepare of adding back to the target.

Dim tdf       As TableDef
Dim db As dao.Database, rst As Recordset
Dim I As Long

Set db = CurrentDb
Set tdf = db.TableDefs(Output_Table)
On Error Resume Next
tdf.Fields.Append tdf.CreateField("delete_matching", 10)     ' dbText
On Error GoTo 0

'Step 1 - Mark all of the matching rows in the Target Table that are matched.
'UPDATE bank_account INNER JOIN imported_table ON (bank_account.bank_account_id = imported_table.bank_account_id) SET bank_account.delete_matching = 'Y';
strSql = "UPDATE " & Output_Table & " INNER JOIN " & imported_table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then _
    strSql = strSql & _
       "(" & imported_table & "." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET " & Output_Table & ".delete_matching = 'Y';"
DebugPrint ("(Delete_Matching_From_Target) Step #1 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all matching rows...


'Step 2 - Save the Create Stamp values from the Target Table.
'UPDATE imported_table INNER JOIN bank_account ON imported_table.bank_account_id = bank_account.bank_account_id
'SET imported_table.create_date_time = [bank_account].[create_date_time], imported_table.create_user = [bank_account].[create_user], imported_table.create_program = [bank_account].[create_program], imported_table.create_file = [bank_account].[create_file];
strSql = "UPDATE " & imported_table & " INNER JOIN " & Output_Table & " ON "
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then _
    strSql = strSql & _
       "(" & imported_table & "." & Br(aField_Name_Output(I)) & " = " & Output_Table & "." & Br(aField_Name_Output(I)) & ") AND "
Next I
If Right(strSql, 5) = " AND " Then strSql = Left(strSql, Len(strSql) - 5)
strSql = strSql & " SET "
strSql = strSql & imported_table & ".create_date_time = [" & Output_Table & "].[create_date_time], "
strSql = strSql & imported_table & ".create_user = [" & Output_Table & "].[create_user], "
strSql = strSql & imported_table & ".create_program = [" & Output_Table & "].[create_program], "
strSql = strSql & imported_table & ".create_file = [" & Output_Table & "].[create_file];"
DebugPrint ("(Delete_Matching_From_Target) Step #2 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all matching rows...

'Step 3 - Actually delete the records from the target table. (Will be added back from imported_table)
strSql = "DELETE " & Output_Table & ".delete_matching FROM " & Output_Table & " WHERE ((" & Output_Table & ".delete_matching)='Y');"
DebugPrint ("(Delete_Matching_From_Target) Step #3 - " & strSql)
DoCmd.RunSQL (strSql)

'Step 4 - Actually delete the records from the target table. (Will be added back from imported_table)
strSql = "UPDATE " & imported_table & " SET " & imported_table & ".matched_target_table = 'N';"
DebugPrint ("(Delete_Matching_From_Target) Step #4 - " & strSql)
DoCmd.RunSQL (strSql)   '  Mark all rows...

tdf.Fields.Delete ("delete_matching") ' Remove the field column from the table.
Set tdf = Nothing
Set db = Nothing

End Sub






' Verify data found in the specification.
Private Function Specification_Is_Valid(Worksheet_Name As String, Output_Table As String)

Dim I As Long, J As Long
Dim ErrorMsg  As String

Specification_Is_Valid = True

' Check to see if all table names are valid for this table.
For I = 1 To numOfSpecs
  If aField_Name_Output(I) <> "" And Not FieldExists(aField_Name_Output(I), Output_Table) Then
    ErrorMsg = "Field name-'" & aField_Name_Output(I) & "' for Excel-'" & aExcel_Heading_Text(I) & "' DOES NOT EXIST in Table-'" & _
            Output_Table & "' specification." & Chr(13) & Chr(13) & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
    Specification_Is_Valid = False
  End If
Next I


' aActive_flag_name and aMark_active_flag work together.  If one is filled in then both are required.
If aActive_flag_name = "" And aMark_active_flag = "" Then GoTo Skip_Active_Validation
' Check to see if the active_flag_name is valid for this table.
If Not FieldExists(aActive_flag_name, Output_Table) Then
    ErrorMsg = "Field name-'" & aActive_flag_name & "' defining an Active Flag, DOES NOT EXIST in Table-'" & _
            Output_Table & "' specification." & Chr(13) & Chr(13) & _
            "A valid field name is required when mark_active_flag='" & aMark_active_flag & "' is specified. " & Chr(13) & Chr(13) & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
    Specification_Is_Valid = False
End If

If (aMark_active_flag <> "Changed") And (aMark_active_flag <> "New") And (aMark_active_flag <> "Imported") Then
    ErrorMsg = "mark_active_flag='" & aMark_active_flag & "' is not valid value for Table-'" & _
            Output_Table & "' in import specification." & Chr(13) & Chr(13) & _
            "A valid mark_active_flag is required when active_flag_name-'" & aActive_flag_name & "' is filled in." & Chr(13) & Chr(13) & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
    Specification_Is_Valid = False
End If

Skip_Active_Validation:

' Now check to make sure that key fields are valid.
J = 0
For I = 1 To numOfSpecs
  If aDup_Key_Field(I) Then J = J + 1
Next I
If J = 0 Then
  ErrorMsg = Output_Table & " import specification has no Key Fields marked.  Correct this and restart the import." & Chr(13) & Chr(13) & _
            "Import process will be terminated."
  MsgBox (ErrorMsg)
  Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  Specification_Is_Valid = False
End If

For I = 1 To numOfSpecs
  If aDup_Key_Field(I) And aField_Name_Output(I) = "" Then
    ErrorMsg = "Key Field is marked for " & aExcel_Heading_Text(I) & " and Table-" & Output_Table & ", but no field name is given." & Chr(13) & Chr(13) & _
            "Correct this and restart the import." & Chr(13) & Chr(13) & _
            "Import process will be terminated."
    MsgBox (ErrorMsg)
    Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
  Specification_Is_Valid = False
  End If
Next I

End Function


Private Function Find_Row_Number_Column(prepedFile As String, wrkSheetName As String)

Dim ExcelApp As Object
Dim ExcelBook As Object
Dim ExcelSheet As Object
Dim I As Long, J As Long, ColumnNumber
Dim holdColumnHeading    As String

Dim excelColumns() As String  '  Excel column letters array......
Call Init_Excel_Names(excelColumns)
   
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(prepedFile)

If Right(HoldFileName, 4) = ".csv" Then
   Set ExcelSheet = ExcelBook.Sheets(1)
Else
   Set ExcelSheet = ExcelBook.Sheets(wrkSheetName)
End If

ExcelSheet.Activate
For I = 1 To 150
   ColumnNumber = excelColumns(I) & "1"  ' Construct the cell id for proper row 1 heading....
   holdColumnHeading = ExcelSheet.Range(ColumnNumber).Value
   If holdColumnHeading = "Row_Number" Then
     Find_Row_Number_Column = I
    ' GoTo leaveFunction
   End If
Next I

leaveFunction:
    ExcelBook.Saved = True  ' Avoid the user message when closing Excel Workbook
    ExcelBook.Close
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    Set ExcelSheet = Nothing
    Exit Function

End Function


Private Function Br(sqlName As String)  ' Function used to insert brackets around SQL field names.
Br = "[" & sqlName & "]"
End Function

Private Sub DebugPrint(msg As String)
If aPrintDebugLog Then Debug.Print (msg)
End Sub

Public Sub Remove_Import_Recipe(ByVal ToName As String)

Dim strSql    As String

' Delete/Remove the Spec3 records from the recipe.
strSql = "DELETE import_spec3_fields.* " _
       & "FROM (import_spec1_file_name INNER JOIN import_spec2_worksheet_name ON import_spec1_file_name.ID = import_spec2_worksheet_name.input_file_name_ID) INNER JOIN import_spec3_fields ON import_spec2_worksheet_name.ID2 = import_spec3_fields.worksheet_name_id " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & ToName & """));"
DebugPrint ("Remove_Import_Recipe- " & strSql)
'DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
'DoCmd.SetWarnings True

' Delete/Remove the Spec2 records from the recipe.
strSql = "DELETE import_spec2_worksheet_name.* " _
       & "FROM import_spec1_file_name INNER JOIN import_spec2_worksheet_name ON import_spec1_file_name.ID = import_spec2_worksheet_name.input_file_name_ID " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & ToName & """));"
DebugPrint ("Remove_Import_Recipe- " & strSql)
'DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
'DoCmd.SetWarnings True

' Delete/Remove the Spec1 records from the recipe.
strSql = "DELETE import_spec1_file_name.* " _
       & "FROM import_spec1_file_name " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & ToName & """));"
DebugPrint ("Remove_Import_Recipe- " & strSql)
'DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
'DoCmd.SetWarnings True

End Sub

Public Sub Clone_Import_Recipe(Optional ByVal FromName As String = "UnionStateTemp", _
                               Optional ByVal ToName As String = "NewName2", _
                               Optional ByVal ToInputFileName As String = "MyInputFile.xlsx", _
                               Optional ByVal ErrHeader As String = "Clone_Import_Recipe routine..............")

' This Routine will clone make a clone recipe that can then be modified as needed.
Dim rst   As Recordset
Dim I As Long, J As Long
Dim strSql   As String
Dim MsgResponse As Long

Dim db        As dao.Database:  Set db = CurrentDb
Dim tdf       As dao.TableDef

' First Check to see if Spec1 FromName exists.
strSql = "SELECT import_spec1_file_name.ID, import_spec1_file_name.spec_name FROM import_spec1_file_name " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & FromName & """));"

DebugPrint ("ImportWorkSheetSpecs - " & strSql)
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
   rst.Close
   Set rst = Nothing
   ErrorMsg = "Clone_Import_Recipe error - Spec1 for """ & FromName & """ was not found." & Chr(13) _
        & Chr(13) & "Clone Process will be ABORTED....."
   Call MsgBox(ErrorMsg, vbOKOnly, ErrHeader)
   Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
   End
End If
Dim HoldOldID   As Long: HoldOldID = rst!Id  ' Get the new ID
rst.Close
Set rst = Nothing

' Second - Check to see if the Spec1 ToName exists and give the user an option to delete it first.
strSql = "SELECT import_spec1_file_name.ID, import_spec1_file_name.spec_name, import_spec1_file_name.input_file_name FROM import_spec1_file_name " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & ToName & """));"

DebugPrint ("ImportWorkSheetSpecs - " & strSql)
Dim ErrorMsg  As String
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount <> 0 Then
   ErrorMsg = "Clone_Import_Recipe error - Spec1 for """ & ToName & """/""" & rst!Input_File_Name & """ was found (NOT Expected)." & Chr(13) _
        & Chr(13) & "Do want to replace?" _
        & Chr(13) & "NO - Clone Process will be ABORTED....."
   MsgResponse = MsgBox(ErrorMsg, vbYesNo, ErrHeader)
   Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
   If MsgResponse = vbNo Then End
   Call Remove_Import_Recipe(ToName)
End If
rst.Close
Set rst = Nothing


'First clone the Spec1 record....

strSql = "INSERT INTO import_spec1_file_name ( spec_name, input_file_name, active, spec1_recipe_desc, accept_changes, allow_change_2_blank, reject_err_file, reject_err_rows, save_date_changed ) " _
       & "SELECT """ & ToName & """ AS Expr1, """ & ToInputFileName & """ AS Expr2, import_spec1_file_name.active, import_spec1_file_name.spec1_recipe_desc, import_spec1_file_name.accept_changes, import_spec1_file_name.allow_change_2_blank, import_spec1_file_name.reject_err_file, import_spec1_file_name.reject_err_rows, import_spec1_file_name.save_date_changed " _
       & "FROM import_spec1_file_name " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & FromName & """));"

DebugPrint ("ImportWorkSheetSpecs - " & strSql)
DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True

' Next read the record just added to get the ID that was assigned by the database.
strSql = "SELECT import_spec1_file_name.ID, import_spec1_file_name.spec_name FROM import_spec1_file_name " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & ToName & """));"
Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then
   rst.Close
   Set rst = Nothing
   ErrorMsg = "Clone_Import_Recipe error - Spec1 for """ & ToName & """ new record just added, was not found." & Chr(13) _
        & Chr(13) & "Clone Process will be ABORTED....."
   Call MsgBox(ErrorMsg, vbOKOnly, ErrHeader)
   Debug.Print (Chr(13) & "****" & ErrorMsg & Chr(13))
   End
End If
Dim HoldNewID   As Long: HoldNewID = rst!Id  ' Get the new ID
rst.Close
Set rst = Nothing

Set tdf = db.TableDefs("import_spec2_worksheet_name")
If Not FieldExists("oldID", "import_spec2_worksheet_name") Then _
  tdf.Fields.Append tdf.CreateField("oldID", 7)   ' dbDouble 7
Set tdf = Nothing

'  Next, clone ALL Spec2 records.......
strSql = "INSERT INTO import_spec2_worksheet_name ( input_file_name_ID, work_sheet_name, output_table_name, active_flag_name, mark_active_flag, unique_import_key_value, active2, accept_changes2, allow_change_2_blank2, reject_err_file2, reject_err_rows2, save_date_changed2, oldID ) " _
       & "SELECT " & HoldNewID & " AS Expr1, import_spec2_worksheet_name.work_sheet_name, import_spec2_worksheet_name.output_table_name, import_spec2_worksheet_name.active_flag_name, import_spec2_worksheet_name.mark_active_flag, import_spec2_worksheet_name.unique_import_key_value, import_spec2_worksheet_name.active2, import_spec2_worksheet_name.accept_changes2, import_spec2_worksheet_name.allow_change_2_blank2, import_spec2_worksheet_name.reject_err_file2, import_spec2_worksheet_name.reject_err_rows2, import_spec2_worksheet_name.save_date_changed2, import_spec2_worksheet_name.ID2 " _
       & "FROM import_spec2_worksheet_name WHERE (((import_spec2_worksheet_name.input_file_name_ID)=" & HoldOldID & "));"


DebugPrint ("ImportWorkSheetSpecs - " & strSql)
DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
DoCmd.SetWarnings True

'  Now SELECT the ALL of new Spec2 records just added and then clone all Spec3 records for each one.
strSql = "INSERT INTO import_spec3_fields ( worksheet_name_id, key_field, excel_column_number, excel_heading_text, field_name_output, active3, accept_changes3, allow_change_2_blank3, reject_err_file3, reject_err_rows3, save_date_changed3 ) " _
       & "SELECT import_spec2_worksheet_name.ID2, import_spec3_fields.key_field, import_spec3_fields.excel_column_number, import_spec3_fields.excel_heading_text, import_spec3_fields.field_name_output, import_spec3_fields.active3, import_spec3_fields.accept_changes3, import_spec3_fields.allow_change_2_blank3, import_spec3_fields.reject_err_file3, import_spec3_fields.reject_err_rows3, import_spec3_fields.save_date_changed3 " _
       & "FROM import_spec1_file_name INNER JOIN (import_spec2_worksheet_name INNER JOIN import_spec3_fields ON import_spec2_worksheet_name.oldID = import_spec3_fields.worksheet_name_id) ON import_spec1_file_name.ID = import_spec2_worksheet_name.input_file_name_ID " _
       & "WHERE (((import_spec1_file_name.spec_name)=""" & ToName & """));"
DebugPrint ("ImportWorkSheetSpecs - " & strSql)
'DoCmd.SetWarnings False
DoCmd.RunSQL (strSql)
'DoCmd.SetWarnings True

Call Remove_Table_Field("oldID", "import_spec2_worksheet_name")

End Sub

