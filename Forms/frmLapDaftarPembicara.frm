VERSION 5.00
Begin VB.Form frmLapDaftarPembicara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Pembicara"
   ClientHeight    =   885
   ClientLeft      =   2385
   ClientTop       =   4560
   ClientWidth     =   4725
   Icon            =   "frmLapDaftarPembicara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4725
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   1
      Left            =   2250
      TabIndex        =   1
      Top             =   450
      Width           =   1905
   End
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Top             =   90
      Width           =   1905
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   0
      Left            =   4185
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   330
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   1
      Left            =   4185
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   135
      Width           =   2040
   End
   Begin VB.Label lblField 
      Caption         =   "Sampai No. Arsip"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   495
      Width           =   2040
   End
End
Attribute VB_Name = "frmLapDaftarPembicara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IReport

Private oReport As CReport
Private sSortOrder As String
Private sReportFileName As String
Private bSaveSetting As Boolean

Public Property Get SaveSettings() As Boolean
    SaveSettings = bSaveSetting
End Property

Public Property Let SaveSettings(ByVal RHS As Boolean)
    bSaveSetting = RHS
End Property

Private Sub cmdPrompt_Click(Index As Integer)
    Dim sSQL As String, PromptTitle As String, ColumnsTitle As Variant
    
    PromptTitle = "Prompt Pembicara"
    Select Case UCase$(sSortOrder)
        Case "NO. ARSIP"
            sSQL = "SELECT NoArsip, NamaLengkap FROM Pembicara"
            ColumnsTitle = Array("No. Arsip", "Nama Lengkap")
        Case "NAMA LENGKAP"
            sSQL = "SELECT NamaLengkap, NoArsip FROM Pembicara"
            ColumnsTitle = Array("Nama Lengkap", "No. Arsip")
        Case "NAMA PANGGILAN"
            sSQL = "SELECT NamaPanggilan, NoArsip FROM Pembicara"
            ColumnsTitle = Array("Nama Panggilan", "No. Arsip")
    End Select
    If oDialog.ShowPrompt(ConnectString, sSQL, PromptTitle, ColumnsTitle, txtField(Index).Text) = DialogAnswerOK Then
        txtField(Index).Text = oDialog.ColumnValue(0)
        txtField(Index).SetFocus
    End If
End Sub

Public Sub IReport_SetSortOrder()
    Dim Choices(5) As String
    
    Choices(0) = "No. Arsip"
    Choices(1) = "Nama Lengkap"
    Choices(2) = "Nama Panggilan"
    
    If oDialog.ShowChoice(Choices, BLANK) = DialogAnswerOK Then
        sSortOrder = UCase$(oDialog.SelectedChoice)
    End If
    
    Call SortOrderChanges
End Sub

Private Sub SortOrderChanges()
    If sSortOrder = BLANK Then sSortOrder = "NO. ARSIP"
    
    lblField(0).Caption = "Dari " & StrConv(sSortOrder, vbProperCase)
    lblField(1).Caption = "Sampai " & StrConv(sSortOrder, vbProperCase)
End Sub

Public Sub IReport_PreviewReport()
    Dim bSuccess As Boolean
    Dim sSelection As String

    '// Beginning of template
    If Not DataIsValid Then
        Exit Sub
    End If
    '//
    
    
    With oReport
        '// Beginning of template
        If Not .OpenReport(sReportFileName) Then
            MsgBox "Cannot found report file or no default printer installed.", vbExclamation, Caption
            Exit Sub
        End If
        '//
        
        Select Case UCase$(sSortOrder)
            Case "NO. ALUMNI", BLANK
                sSelection = "{Pembicara.NoArsip} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "NAMA LENGKAP"
                sSelection = "{Pembicara.NamaLengkap} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "NAMA PANGGILAN"
                sSelection = "{Pembicara.NamaPanggilan} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
        End Select
        bSuccess = .SetRecordselection(sSelection)
        
        bSuccess = .SetDatabaseLocation(0, MainDB.Properties("Data Source").Value)
        bSuccess = .ShowReportToWindow(frmMain, Me.Caption)
    End With

    
    '//
End Sub

Private Function DataIsValid() As Boolean
    '// Do input validation here...
    
    If Trim$(txtField(1).Text) = BLANK Then txtField(1).Text = "zzz"
    
    sReportFileName = App.Path & "\Reports\DaftarPembicara.rpt"

    '// Beginning of template
    DataIsValid = True
    '//
End Function

Private Sub Form_Activate()
    '// Beginning of template
    Call ToggleReportToolbar(True)
    '// End of template
    
    'Your code here...
End Sub

Private Sub Form_Deactivate()
    '// Beginning of template
    Call ToggleReportToolbar(False)
    '// End of template
    
    'Your code here...
End Sub

Private Sub Form_GotFocus()
    '// Beginning of template
    Call ToggleReportToolbar(True)
    '// End of template
    
    'Your code here...
End Sub

Private Sub Form_Load()
    '// Beginning of Template
    Call LoadPosition(Me, INIPath)
    Set oReport = New CReport
    
    sSortOrder = UCase$(ReadINIFile(Me.Caption, "SortOrder", BLANK, INIPath))
    Call SortOrderChanges
    '//
End Sub

Private Sub Form_LostFocus()
    '// Beginning of template
    Call ToggleReportToolbar(False)
    '// End of template
    
    'Your code here...
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '// Beginning of template
    Call SavePosition(Me, INIPath)
    
    If SaveSettings Then
        Call WriteINIFile(Me.Caption, "SortOrder", sSortOrder, INIPath)
    End If
    
    Call ToggleReportToolbar(False)
    
    Set oReport = Nothing
    '//
End Sub

Private Sub txtField_GotFocus(Index As Integer)
    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
End Sub

