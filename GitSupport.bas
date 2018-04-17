Attribute VB_Name = "GitSupport"
Option Compare Database
Option Explicit

Dim ThisModuleName   As String


Public Function Gexp()  '  Git Export
Dim I  As Long

'  First locate the GIT folder that is located in the same folder as the Office project.
Dim GitProjectFolderName  As String
GitProjectFolderName = Application.CurrentProject.Name
I = InStrRev(GitProjectFolderName, ".")
If I > 1 Then GitProjectFolderName = Left(GitProjectFolderName, I - 1)
GitProjectFolderName = Application.CurrentProject.Path & "\" & GitProjectFolderName & "\"

Call MakeThisFolder(GitProjectFolderName)  '  Create folder if it does not exist...
Call ExportSourceFiles(GitProjectFolderName)
Call MsgBox("Git VBA Export to: " & Chr(13) & Chr(13) & GitProjectFolderName & Chr(13) & Chr(13) & "   is COMPLETE...")

End Function

Public Function Gimp()   '  Git Import all modules from the project directory.

ThisModuleName = "GitSupport"  ' Name of this module....

'  First locate the GIT folder that is located in the same folder as the Office project.
Dim GitProjectFolderName  As String
Dim I  As Long
GitProjectFolderName = Application.CurrentProject.Name
I = InStrRev(GitProjectFolderName, ".")
If I > 1 Then GitProjectFolderName = Left(GitProjectFolderName, I - 1)
GitProjectFolderName = Application.CurrentProject.Path & "\" & GitProjectFolderName & "\"
  
Call ImportSourceFiles
Call MsgBox("Git VBA Import to: " & Chr(13) & Chr(13) & GitProjectFolderName & Chr(13) & Chr(13) & "   is COMPLETE...")
End Function


Private Sub ImportSourceFiles()

Dim I   As Long
Dim ModuleName   As String

'  First locate the GIT folder that is located in the same folder as the Office project.
Dim GitProjectFolderName  As String
GitProjectFolderName = Application.CurrentProject.Name
I = InStrRev(GitProjectFolderName, ".")
If I > 1 Then GitProjectFolderName = Left(GitProjectFolderName, I - 1)
GitProjectFolderName = Application.CurrentProject.Path & "\" & GitProjectFolderName & "\"

Dim file As String
file = Dir(GitProjectFolderName)
While (file <> vbNullString)
  Call RemoveAModule(file)
  I = InStrRev(file, ".")
  If I > 1 Then ModuleName = Left(file, I - 1)
  If ThisModuleName <> ModuleName Then _
    Application.VBE.ActiveVBProject.VBComponents.Import GitProjectFolderName & file
  file = Dir
Wend

End Sub

Private Sub RemoveAllModules()
Dim project As VBProject
Set project = Application.VBE.ActiveVBProject
 
Dim comp As VBComponent
For Each comp In project.VBComponents
  If Not comp.Name = GitProjectFolderName And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
    project.VBComponents.Remove comp
  End If
Next
End Sub


Private Sub RemoveAModule(ByVal ModuleName As String)
Dim project As VBProject
Dim I   As Long

I = InStrRev(ModuleName, ".")
If I > 1 Then ModuleName = Left(ModuleName, I - 1)

If ModuleName = ThisModuleName Then Exit Sub  '  Don't remove this module.

Set project = Application.VBE.ActiveVBProject
 
Dim comp As VBComponent
For Each comp In project.VBComponents
  If comp.Name = ModuleName And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
    project.VBComponents.Remove comp
    Exit Sub
  End If
Next
End Sub


Public Sub ExportSourceFiles(destPath As String)
 
Dim component As VBComponent
Dim KillFileAndPath  As String
For Each component In Application.VBE.ActiveVBProject.VBComponents
  If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
    KillFileAndPath = destPath & component.Name & ToFileExtension(component.Type)
    '''Call Kill(KillFileAndPath)
    component.Export destPath & component.Name & ToFileExtension(component.Type)
  End If
Next
 
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function

Public Sub MakeThisFolder(FolderAndPath As String)
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
