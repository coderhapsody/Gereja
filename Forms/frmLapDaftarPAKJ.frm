VERSION 5.00
Begin VB.Form frmLapDaftarPAKJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar PAKJ"
   ClientHeight    =   990
   ClientLeft      =   2520
   ClientTop       =   2355
   ClientWidth     =   4545
   Icon            =   "frmLapDaftarPAKJ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4545
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   1
      Left            =   4140
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   450
      Width           =   330
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   0
      Left            =   4140
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   330
   End
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   0
      Left            =   2205
      TabIndex        =   1
      Top             =   90
      Width           =   1905
   End
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   1
      Left            =   2205
      TabIndex        =   0
      Top             =   450
      Width           =   1905
   End
   Begin VB.Label lblField 
      Caption         =   "Sampai No. Arsip"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   495
      Width           =   2040
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   2040
   End
End
Attribute VB_Name = "frmLapDaftarPAKJ"
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

Public Sub IReport_SetSortOrder()
    Call SortOrderChanges
End Sub

Private Sub SortOrderChanges()
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
        
        bSuccess = .SetRecordselection(sSelection)
        
        bSuccess = .SetDatabaseLocation(0, MainDB.Properties("Data Source").Value)
        bSuccess = .ShowReportToWindow(frmMain, Me.Caption)
    End With

    
    '//
End Sub

Private Function DataIsValid() As Boolean
    '// Do input validation here...
    
    If Trim$(txtField(1).Text) = BLANK Then txtField(1).Text = "zzz"
    
    sReportFileName = App.Path & "\Reports\DaftarPAKJ.rpt"

    '// Beginning of template
    DataIsValid = True
    '//
End Function

