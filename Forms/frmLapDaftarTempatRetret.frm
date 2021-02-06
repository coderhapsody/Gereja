VERSION 5.00
Begin VB.Form frmLapDaftarTempatRetret 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Tempat Retret"
   ClientHeight    =   945
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   4695
   Icon            =   "frmLapDaftarTempatRetret.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4695
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   1
      Left            =   2205
      TabIndex        =   3
      Top             =   450
      Width           =   1905
   End
   Begin VB.TextBox txtField 
      Height          =   330
      Index           =   0
      Left            =   2205
      TabIndex        =   2
      Top             =   90
      Width           =   1905
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   0
      Left            =   4140
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   330
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "..."
      Height          =   330
      Index           =   1
      Left            =   4140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   450
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   135
      Width           =   2040
   End
   Begin VB.Label lblField 
      Caption         =   "Sampai No. Arsip"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   495
      Width           =   2040
   End
End
Attribute VB_Name = "frmLapDaftarTempatRetret"
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
    End With

    
    '//
End Sub

Private Function DataIsValid() As Boolean
    '// Do input validation here...
    
    If Trim$(txtField(1).Text) = BLANK Then txtField(1).Text = "zzz"
    
    sReportFileName = App.Path & "\Reports\DaftarTempatRetret.rpt"

    '// Beginning of template
    DataIsValid = True
    '//
End Function
