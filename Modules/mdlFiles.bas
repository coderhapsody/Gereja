Attribute VB_Name = "mdlFiles"
'------------------------------------------------------------------------------------------------------------------
'
'  mdlFiles.bas
'
'  Contains wrapped functions to file access.
'
'
'  (C)Paulus Iman, November 2003-Januari 2005
'  Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public Function XCopyFile(ByVal SourceFile As String, ByVal DestFile As String, ByVal ReplaceOldFile As Boolean) As Boolean
    XCopyFile = CBool(CopyFile(SourceFile, DestFile, CLng(CInt(ReplaceOldFile) * -1)))
End Function

Public Function ReadINIFile(ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal INIPath As String) As String
    Dim sNilKembali As String, lSukses As Long
    
    sNilKembali = String$(256, 0)
    lSukses = GetPrivateProfileString(ByVal Section, ByVal Key, ByVal Default, _
              ByVal sNilKembali, Len(sNilKembali), ByVal INIPath)
    If lSukses = 0 Then
        ReadINIFile = ""
    Else
        ReadINIFile = Left$(sNilKembali, InStr(sNilKembali, Chr(0)) - 1)
    End If
End Function

Public Function WriteINIFile(ByVal Section As String, ByVal Key As String, ByVal Value As String, ByVal INIPath As String) As Boolean
    WriteINIFile = CBool(WritePrivateProfileString(Section, Key, Value, INIPath))
End Function


