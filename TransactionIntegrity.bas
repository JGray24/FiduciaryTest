Attribute VB_Name = "TransactionIntegrity"
Option Compare Database
Option Explicit

Public Function transactionIntegrityCheck(Optional ByVal pHeader As String) As String

' This routine will validate various parts of the QBO file received from the bank.

'1) Bank account is found in the "bank_account" database table.
'    Call confirmValidAccount
'2) Routing number matches bank routing number found the "bank" database table.
'    Call confirmBankRouting

'3) Running balance has integrity and is in balance with itself.
    Call ensureRunningBalanceIntegrity(pHeader)

'4) All "account_designator" fields are filled in and match with the "bank_account" database table.
'    Call ensureBankAccountsProperlyLinked
    
'4a) All fields have a valid veteran case number.
'   Call ensure_ValidCaseNumber

'5) All "account_designator" fields found in the "transaction" table have integrity with the "bank_account", before adding
'   new records, make sure that the bank account tables have not been changed to mis-match what is in the QBO table.




End Function


Private Sub ensureRunningBalanceIntegrity(ByVal pHeader As String)

'3) Running balance has integrity and is in balance with itself.
If pHeader = "" Then pHeader = "General Integrity Check "
If Right(pHeader, 1) <> " " Then pHeader = pHeader & " "
pHeader = pHeader & "***************************************"

Dim strSql              As String
Dim rst                 As Recordset
Dim runningBalance      As Double
Dim missingTransactions As Double
Dim thisKey             As String, lastKey As String: lastKey = "???"
Dim holdPrevDate        As Date

strSql = "UPDATE transactions SET transactions.missing_trans = 0;"
DoCmd.RunSQL (strSql)
strSql = "UPDATE transactions SET transactions.prev_date_posted = NULL;"
DoCmd.RunSQL (strSql)

strSql = "SELECT transactions.account_designator, transactions.posted_date, transactions.serial_number, transactions.qbo_row_number, transactions.amount, transactions.running_balance, transactions.missing_trans, transactions.prev_date_posted " _
       & "FROM transactions " _
       & "ORDER BY transactions.account_designator, transactions.posted_date, transactions.qbo_row_number;"
Debug.Print ("Integrity Check 3-a) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql)  ' Open with intent to edit
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all account_designator(s) are filled in.....
Dim displayMsg   As String
Do
  thisKey = rst!account_designator
  If thisKey <> lastKey Then runningBalance = rst!running_balance - rst!amount
  lastKey = thisKey
  runningBalance = runningBalance + rst!amount
  
  missingTransactions = runningBalance - rst!running_balance
  If missingTransactions <> 0 Then
    rst.Edit
    rst!missing_trans = missingTransactions * -1
    rst!prev_date_posted = holdPrevDate
    runningBalance = runningBalance + rst!missing_trans
    rst.Update
  End If
  
  holdPrevDate = Nz(rst!posted_date, 0)
  rst.MoveNext
  If rst.EOF Then GoTo Finished_Do_Loop2
Loop
Finished_Do_Loop2:
rst.Close

' Now check to see if any transactions are missing.....
strSql = "SELECT transactions.account_designator, transactions.prev_date_posted, transactions.missing_trans, transactions.posted_date, transactions.amount, transactions.payee " _
       & "FROM transactions " _
       & "WHERE (((transactions.missing_trans)<>0));"
Debug.Print ("Integrity Check 3-b) " & strSql)

Set rst = Application.CurrentDb.OpenRecordset(strSql, dbReadOnly)
If rst.RecordCount = 0 Then Exit Sub
'  Now loop thru the query results and confirm that all account_designator(s) are filled in.....

displayMsg = "Integrity Check found transaction(s) missing/out of balance from MAIN ""transactions"" table. " & Chr(13) & Chr(13) & _
             "Missing/Out of Balance  Errors found:" & Chr(13) & Chr(13)

Do
  displayMsg = displayMsg & "Account - " & Nz(rst!account_designator, "") & " / $" & Nz(rst!missing_trans, 0) & " is missing between " & _
               Nz(rst!prev_date_posted, "") & " and " & Nz(rst!posted_date, "") & Chr(13) & Chr(13)
  
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


