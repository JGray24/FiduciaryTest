Attribute VB_Name = "Globals"
Option Compare Database
Option Explicit

' Access global variables definition...
Global GBL_file_name  As Double

Public Sub Init_Globals()
' Access global variable initialization
  GBL_file_name = 0  '  Will be used to increment a file number to store in union_state_temp_qbo table
End Sub
