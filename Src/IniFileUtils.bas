Attribute VB_Name = "IniFileUtils"
Option Explicit

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long
  
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long
  
Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String

  Const MaxStrSize = 255

  Dim Worked As Long
  Dim RetStr As String * MaxStrSize
  Dim StrSize As Long
  Dim iNoOfCharInIni As Integer
  Dim sIniString As String
  Dim sProfileString As String

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(MaxStrSize)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, gIniFilePath)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
  
End Function

Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
  
  Dim Worked As Long
  Dim iNoOfCharInIni As Integer
  Dim sIniString As String

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Sect, Keyname, Wstr, gIniFilePath)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteIniFileString = sIniString
  End If
  
End Function

Function GetBooleanSetting(Key As String) As Boolean

  Dim Value As String
  
  Value = ReadIniFileString(MainSection, Key)
  GetBooleanSetting = StrComp(Value, "True", vbTextCompare) = 0

End Function
