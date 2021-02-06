Attribute VB_Name = "mdlGeneralFunctions"
'------------------------------------------------------------------------------------------------------------------
'
'  mdlGeneralFunctions.bas
'
'  Contains wrapped commonly used functions. Please read the documentation at the above of each functions/subs.
'
'
'  (C)Paulus Iman, November 2003-Januari 2005
'  Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------

Option Explicit

'Late binding to StatusBar object in frmMain
Private m_StatusBar As Object

'Late binding to ToolBar object in frmMain
Private m_Toolbar As Object

Public Sub ToggleReportToolbar(ByVal Toggle As Boolean)
    frmMain.picReportToolBar.Visible = Toggle
    frmMain.chkSaveSetting = vbUnchecked
End Sub

Public Function CheckKeyPress(ByVal KeyAscii As Integer) As Integer
    If KeyAscii = vbKeyReturn Then
        CheckKeyPress = 0
        SendKeys "{Tab}"
    Else
        CheckKeyPress = KeyAscii
    End If
End Function

Public Function ToNumeric(ByVal Number As Variant) As Double
    If IsEmpty(Number) Or IsNull(Number) Or Trim(CStr(Number) = BLANK) Then
        ToNumeric = 0
    Else
        ToNumeric = CDbl(Number)
    End If
End Function

Public Function NumberOnly(ByVal KeyAscii As Integer, _
                           Optional ByVal AllowDotComma As Boolean = False, _
                           Optional ByVal AllowPositiveNegative As Boolean = False, _
                           Optional ByVal AllowPercent As Boolean = False) As Integer
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyEscape
             NumberOnly = KeyAscii
        Case Asc("."), Asc(",")
             If Not AllowPercent Then
                If Not AllowDotComma Then
                    NumberOnly = 0
                Else
                    NumberOnly = KeyAscii
                End If
            Else
                If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
                    NumberOnly = KeyAscii
                End If
            End If
        Case Asc("+"), Asc("-")
             If Not AllowPositiveNegative Then
                NumberOnly = 0
             Else
                NumberOnly = KeyAscii
             End If
        Case Else
             NumberOnly = 0
    End Select
End Function

Public Function EmailIsValid(ByVal EmailAddress As String) As Boolean
    ' original by Brad Murray
    ' optimized by Rob Hofker, email: rob@eurocamp.nl,
     '23 august 2000
    
    Dim sInvalidChars As String
    Dim bTemp As Boolean
    Dim i As Integer
    Dim sTemp As String
    Dim sEmail As String
    
    sEmail = EmailAddress

    ' Disallowed characters
    sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< "

    ' Check that there is at least one '@'
    bTemp = InStr(sEmail, "@") <= 0
    If bTemp Then GoTo exit_function

    ' Check that there is at least one '.'
    bTemp = InStr(sEmail, ".") <= 0
    If bTemp Then GoTo exit_function

    ' and that the length is at least six (a@a.ca)
    bTemp = Len(sEmail) < 6
    If bTemp Then GoTo exit_function

    ' Check that there is only one '@'
    i = InStr(sEmail, "@")
    sTemp = Mid$(sEmail, i + 1)
    bTemp = InStr(sTemp, "@") > 0
    
    If bTemp Then GoTo exit_function
    'extra checks
    ' AFTER '@' space is not allowed
    bTemp = InStr(sTemp, " ") > 0
    If bTemp Then GoTo exit_function

    ' Check that there is one dot AFTER '@'
    bTemp = InStr(sTemp, ".") = 0
    If bTemp Then GoTo exit_function
    
    ' Check if there's a quote (")
    bTemp = InStr(sEmail, Chr(34)) > 0
    If bTemp Then GoTo exit_function
    
        
    ' Check if there's any other disallowed chars
    ' optimize a little if sEmail longer than sInvalidChars
    ' check the other way around
    If Len(sEmail) > Len(sInvalidChars) Then
        For i = 1 To Len(sInvalidChars)
            If InStr(sEmail, Mid$(sInvalidChars, i, 1)) > 0 _
                  Then bTemp = True
            If bTemp Then Exit For
        Next
    Else
        For i = 1 To Len(sEmail)
            If InStr(sInvalidChars, Mid$(sEmail, i, 1)) > 0 _
                   Then bTemp = True
            If bTemp Then Exit For
        Next
    End If
    If bTemp Then GoTo exit_function
    
    ' extra check
    ' no two consecutive dots
    bTemp = InStr(sEmail, "..") > 0
    If bTemp Then GoTo exit_function
    
exit_function:
    ' if any of the above are true, invalid e-mail
    EmailIsValid = Not bTemp
End Function

Public Property Get NumberFormat(Optional ByVal Digits As Byte = 2)
    Dim LeadingZeros As String
    
    For iLoop = 1 To Digits
        LeadingZeros = LeadingZeros & "0"
    Next
    
    If Digits > 0 Then
        NumberFormat = "###,##0." & LeadingZeros
    ElseIf Digits = 0 Then
        NumberFormat = "###,###"
    End If
    
End Property

Public Sub InitializeMainControls(MainToolBar As Object, StatusBar As Object)
    Set m_Toolbar = MainToolBar
    Set m_StatusBar = StatusBar
End Sub

Public Sub ShowStatusBar(ByVal Message As String)
    With m_StatusBar
        If UCase$(Message) = "RESET" Then
            .Panels(1).Text = "Ready"
            Exit Sub
        End If
        .Panels(1).Text = Message
        DoEvents
    End With
End Sub

Public Sub AdjustToolbarButton(ByVal AddNewButton As Boolean, _
                               ByVal SaveButton As Boolean, _
                               ByVal DeleteButton As Boolean, _
                               ByVal CancelButton As Boolean, _
                               Optional ByVal PrintButton As Boolean = True, _
                               Optional ByVal RefreshButton As Boolean = True)
    With m_Toolbar
        .Buttons("ADDNEW").enabled = AddNewButton
        .Buttons("SAVE").enabled = SaveButton
        .Buttons("DELETE").enabled = DeleteButton
        .Buttons("CANCEL").enabled = CancelButton
        .Buttons("PRINT").enabled = PrintButton
        .Buttons("REFRESH").enabled = RefreshButton
        DoEvents
    End With
End Sub
                   
Public Sub ToggleToolbarButtons(ByVal Toggle As Boolean)
    Call AdjustToolbarButton(Toggle, Toggle, Toggle, Toggle, Toggle, Toggle)
End Sub

Public Sub DisableToolbarButtons()
    Call ToggleToolbarButtons(False)
End Sub

Public Sub EnableToolbarButtons()
    Call ToggleToolbarButtons(True)
End Sub

Public Sub SavePosition(ActiveForm As Object, ByVal INIPath As String)
    Call WriteINIFile(ActiveForm.Caption, "Left", ActiveForm.Left, INIPath)
    Call WriteINIFile(ActiveForm.Caption, "Top", ActiveForm.Top, INIPath)
End Sub

Public Sub LoadPosition(ActiveForm As Object, ByVal INIPath As String)
    ActiveForm.Left = CSng(ReadINIFile(ActiveForm.Caption, "Left", ActiveForm.Left, INIPath))
    ActiveForm.Top = CSng(ReadINIFile(ActiveForm.Caption, "Top", ActiveForm.Top, INIPath))
    DoEvents
End Sub
