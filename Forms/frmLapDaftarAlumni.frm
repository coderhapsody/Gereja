VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLapDaftarAlumni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Alumni"
   ClientHeight    =   915
   ClientLeft      =   2370
   ClientTop       =   3165
   ClientWidth     =   4725
   Icon            =   "frmLapDaftarAlumni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4725
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   330
      Index           =   1
      Left            =   2250
      TabIndex        =   7
      Top             =   450
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      _Version        =   393216
      Format          =   50528257
      CurrentDate     =   38353
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   330
      Index           =   0
      Left            =   2250
      TabIndex        =   6
      Top             =   90
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      _Version        =   393216
      Format          =   50528257
      CurrentDate     =   38353
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   1
      Left            =   4230
      TabIndex        =   5
      Top             =   450
      Width           =   330
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   0
      Left            =   4230
      TabIndex        =   4
      Top             =   90
      Width           =   330
   End
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   1
      Left            =   2250
      TabIndex        =   3
      Top             =   450
      Width           =   1950
   End
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   0
      Left            =   2250
      TabIndex        =   1
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label lblField 
      Caption         =   "Sampai No. Alumni"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   495
      Width           =   2040
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Alumni"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2040
   End
End
Attribute VB_Name = "frmLapDaftarAlumni"
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
    
    PromptTitle = "Prompt Alumni"
    Select Case UCase$(sSortOrder)
        Case "NO. ALUMNI"
            sSQL = "SELECT NoAlumni, NamaLengkap FROM Alumni"
            ColumnsTitle = Array("No. Alumni", "Nama Lengkap")
        Case "NAMA LENGKAP"
            sSQL = "SELECT NamaLengkap, NoAlumni FROM Alumni"
            ColumnsTitle = Array("Nama Lengkap", "No. Alumni")
        Case "NAMA PANGGILAN"
            sSQL = "SELECT NamaPanggilan, NoAlumni FROM Alumni"
            ColumnsTitle = Array("Nama Panggilan", "No. Alumni")
        Case "ANGKATAN"
            sSQL = "SELECT DISTINCT Angkatan FROM Alumni ORDER BY Angkatan"
            ColumnsTitle = Array("Angkatan")
        Case "JURUSAN"
            sSQL = "SELECT DISTINCT Jurusan FROM Alumni ORDER BY Jurusan"
            ColumnsTitle = Array("Jurusan")
    End Select
    If oDialog.ShowPrompt(ConnectString, sSQL, PromptTitle, ColumnsTitle, txtField(Index).Text) = DialogAnswerOK Then
        txtField(Index).Text = oDialog.ColumnValue(0)
        txtField(Index).SetFocus
    End If
End Sub

Public Sub IReport_SetSortOrder()
    Dim Choices(5) As String
    
    Choices(0) = "No. Alumni"
    Choices(1) = "Nama Lengkap"
    Choices(2) = "Nama Panggilan"
    Choices(3) = "Angkatan"
    Choices(4) = "Jurusan"
    Choices(5) = "Tanggal Lahir"
    
    If oDialog.ShowChoice(Choices, BLANK) = DialogAnswerOK Then
        sSortOrder = UCase$(oDialog.SelectedChoice)
    End If
    
    Call SortOrderChanges
End Sub

Private Sub SortOrderChanges()
    If sSortOrder = BLANK Then sSortOrder = "NO. ALUMNI"
    
    If UCase$(sSortOrder) = "TANGGAL LAHIR" Then
        DTPicker(0).Visible = True
        DTPicker(1).Visible = True
        cmdPrompt(0).Visible = False
        cmdPrompt(1).Visible = False
        txtField(0).Visible = False
        txtField(1).Visible = False
    Else
        DTPicker(0).Visible = False
        DTPicker(1).Visible = False
        cmdPrompt(0).Visible = True
        cmdPrompt(1).Visible = True
        txtField(0).Visible = True
        txtField(1).Visible = True
    End If
    
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
                sSelection = "{Alumni.NoAlumni} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "NAMA LENGKAP"
                sSelection = "{Alumni.NamaLengkap} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "NAMA PANGGILAN"
                sSelection = "{Alumni.NamaPanggilan} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "ANGKATAN"
                sSelection = "{Alumni.Angkatan} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "JURUSAN"
                sSelection = "{Alumni.Jurusan} IN '" & txtField(0).Text & "' TO '" & txtField(1).Text & "'"
            Case "TANGGAL LAHIR"
                sSelection = "{Alumni.TanggalLahir} IN #" & DTPicker(0).Value & "# TO #" & DTPicker(1).Value & "#"
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
    
    sReportFileName = App.Path & "\Reports\DaftarAlumni.rpt"

    '// Beginning of template
    DataIsValid = True
    '//
End Function

Private Sub Form_Activate()
    '// Beginning of template
    Call ToggleReportToolbar(True)
    Call DisableToolbarButtons
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
    Call DisableToolbarButtons
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
