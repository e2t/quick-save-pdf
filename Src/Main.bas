Attribute VB_Name = "Main"
Option Explicit

Const IniFileName = "Settings.ini"
Const KeyCloseAfterSave = "CloseAfterSave"
Public Const MainSection = "Main"

Public gFSO As FileSystemObject
Public gIniFilePath As String
Dim swApp As Object

Sub Main()
  
  Dim CurrentDoc As ModelDoc2
  Dim DocPath As String
  
  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType <> swDocDRAWING Then Exit Sub
  Init
  
  DocPath = CurrentDoc.GetPathName
  If DocPath = "" Then
    MsgBox "Сохраните документ.", vbCritical
    Exit Sub
  End If
  SaveAsPDF CurrentDoc, CreateNewName(DocPath)
  If GetBooleanSetting(KeyCloseAfterSave) Then
    swApp.CloseDoc CurrentDoc.GetPathName
  End If
  
End Sub

Function Init() 'hide

  gIniFilePath = gFSO.BuildPath(swApp.GetCurrentMacroPathFolder, IniFileName)
  If Not gFSO.FileExists(gIniFilePath) Then
    CreateDefaultIniFile
  End If

End Function

Sub SaveAsPDF(CurrentDoc As ModelDoc2, NewName As String)

  Dim Errors As swFileSaveError_e
  Dim Warnings  As swFileSaveWarning_e

  CurrentDoc.Extension.SaveAs NewName, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Errors, Warnings

End Sub

Function CreateNewName(Path As String) As String

  CreateNewName = gFSO.BuildPath(gFSO.GetParentFolderName(Path), gFSO.GetBaseName(Path) + ".PDF")

End Function

Function CreateDefaultIniFile() 'hide

  Const DefaultText = "[" + MainSection + "]" + vbNewLine _
    + KeyCloseAfterSave + " = False" + vbNewLine

  Dim objStream As Stream
      
  Set objStream = New Stream
  objStream.Open
  objStream.WriteText DefaultText
  objStream.SaveToFile gIniFilePath

End Function
