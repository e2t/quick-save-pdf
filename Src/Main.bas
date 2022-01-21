Attribute VB_Name = "Main"
Option Explicit

Public gFSO As FileSystemObject
Dim swApp As Object

Sub Main()
  
  Dim CurrentDoc As ModelDoc2
  Dim DocPath As String
  
  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType <> swDocDRAWING Then Exit Sub
  
  DocPath = CurrentDoc.GetPathName
  If DocPath = "" Then
    MsgBox "Сохраните документ.", vbCritical
    Exit Sub
  End If
  SaveAsPDF CurrentDoc, CreateNewName(DocPath)
  
End Sub

Sub SaveAsPDF(CurrentDoc As ModelDoc2, NewName As String)

  Dim Errors As swFileSaveError_e
  Dim Warnings  As swFileSaveWarning_e

  CurrentDoc.Extension.SaveAs NewName, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Errors, Warnings

End Sub

Function CreateNewName(Path As String) As String

  CreateNewName = gFSO.BuildPath(gFSO.GetParentFolderName(Path), gFSO.GetBaseName(Path) + ".PDF")

End Function
