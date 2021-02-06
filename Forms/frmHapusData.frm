VERSION 5.00
Begin VB.Form frmHapusData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hapus Data"
   ClientHeight    =   2955
   ClientLeft      =   2685
   ClientTop       =   3405
   ClientWidth     =   10035
   Icon            =   "frmHapusData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   10035
   Begin VB.CommandButton cmdAction 
      Caption         =   "Batal"
      Height          =   420
      Index           =   1
      Left            =   4950
      TabIndex        =   32
      Top             =   2385
      Width           =   2220
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Hapus"
      Height          =   420
      Index           =   0
      Left            =   2205
      TabIndex        =   31
      Top             =   2385
      Width           =   2220
   End
   Begin VB.CheckBox chkData 
      Caption         =   "Persekutuan Alumni Kristen Jakarta (PAKJ)"
      Height          =   240
      Index           =   4
      Left            =   270
      TabIndex        =   29
      Top             =   1890
      Width           =   3480
   End
   Begin VB.CheckBox chkData 
      Caption         =   "Tempat Retret"
      Height          =   240
      Index           =   3
      Left            =   270
      TabIndex        =   27
      Top             =   1485
      Width           =   3120
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   9
      Left            =   9090
      TabIndex        =   26
      Top             =   1845
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   8
      Left            =   6660
      TabIndex        =   25
      Top             =   1845
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   7
      Left            =   9090
      TabIndex        =   24
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   6
      Left            =   6660
      TabIndex        =   23
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   5
      Left            =   9090
      TabIndex        =   22
      Top             =   1035
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   4
      Left            =   6660
      TabIndex        =   21
      Top             =   1035
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   3
      Left            =   9090
      TabIndex        =   20
      Top             =   630
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   2
      Left            =   6660
      TabIndex        =   19
      Top             =   630
      Width           =   375
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   9
      Left            =   7560
      TabIndex        =   18
      Top             =   1845
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   8
      Left            =   5130
      TabIndex        =   17
      Top             =   1845
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   7
      Left            =   7560
      TabIndex        =   16
      Top             =   1440
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   6
      Left            =   5130
      TabIndex        =   15
      Top             =   1440
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   5
      Left            =   7560
      TabIndex        =   14
      Top             =   1035
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      Left            =   5130
      TabIndex        =   13
      Top             =   1035
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      Left            =   7560
      TabIndex        =   12
      Top             =   630
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   5130
      TabIndex        =   11
      Top             =   630
      Width           =   1500
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   1
      Left            =   9090
      TabIndex        =   10
      Top             =   225
      Width           =   375
   End
   Begin VB.CommandButton cmdPrompt 
      Height          =   330
      Index           =   0
      Left            =   6660
      TabIndex        =   9
      Top             =   225
      Width           =   375
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   7560
      TabIndex        =   8
      Top             =   225
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   5130
      TabIndex        =   7
      Top             =   225
      Width           =   1500
   End
   Begin VB.CheckBox chkData 
      Caption         =   "Pembicara"
      Height          =   240
      Index           =   2
      Left            =   270
      TabIndex        =   2
      Top             =   1080
      Width           =   3120
   End
   Begin VB.CheckBox chkData 
      Caption         =   "Gereja dan Organisasi Kristen Lainnya"
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   675
      Width           =   3120
   End
   Begin VB.CheckBox chkData 
      Caption         =   "Alumni"
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   1770
   End
   Begin VB.Label lblField 
      Caption         =   "s/d"
      Height          =   195
      Index           =   9
      Left            =   7110
      TabIndex        =   36
      Top             =   1935
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "s/d"
      Height          =   195
      Index           =   8
      Left            =   7110
      TabIndex        =   35
      Top             =   1518
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "s/d"
      Height          =   195
      Index           =   7
      Left            =   7110
      TabIndex        =   34
      Top             =   1102
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "s/d"
      Height          =   195
      Index           =   6
      Left            =   7110
      TabIndex        =   33
      Top             =   686
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   4
      Left            =   3915
      TabIndex        =   30
      Top             =   1890
      Width           =   1140
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   3
      Left            =   3915
      TabIndex        =   28
      Top             =   1485
      Width           =   1140
   End
   Begin VB.Label lblField 
      Caption         =   "s/d"
      Height          =   195
      Index           =   5
      Left            =   7110
      TabIndex        =   6
      Top             =   270
      Width           =   330
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   2
      Left            =   3915
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Arsip"
      Height          =   195
      Index           =   1
      Left            =   3915
      TabIndex        =   4
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label lblField 
      Caption         =   "Dari No. Alumni"
      Height          =   195
      Index           =   0
      Left            =   3915
      TabIndex        =   3
      Top             =   270
      Width           =   1140
   End
End
Attribute VB_Name = "frmHapusData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkData_Click(Index As Integer)
    Select Case Index
        Case 0
            If chkData(Index).Value = vbChecked Then
                txtField(0).BackColor = vbWindowBackground
                txtField(1).BackColor = vbWindowBackground
            Else
                txtField(0).BackColor = vbButtonFace
                txtField(1).BackColor = vbButtonFace
            End If
            txtField(0).enabled = chkData(Index).Value = vbChecked
            txtField(1).enabled = chkData(Index).Value = vbChecked
            
        Case 1
            If chkData(Index).Value = vbChecked Then
                txtField(2).BackColor = vbWindowBackground
                txtField(3).BackColor = vbWindowBackground
            Else
                txtField(2).BackColor = vbButtonFace
                txtField(3).BackColor = vbButtonFace
            End If
            txtField(2).enabled = chkData(Index).Value = vbChecked
            txtField(3).enabled = chkData(Index).Value = vbChecked
            
        Case 2
            If chkData(Index).Value = vbChecked Then
                txtField(4).BackColor = vbWindowBackground
                txtField(5).BackColor = vbWindowBackground
            Else
                txtField(4).BackColor = vbButtonFace
                txtField(5).BackColor = vbButtonFace
            End If
            txtField(4).enabled = chkData(Index).Value = vbChecked
            txtField(5).enabled = chkData(Index).Value = vbChecked
            
        Case 3
            If chkData(Index).Value = vbChecked Then
                txtField(6).BackColor = vbWindowBackground
                txtField(7).BackColor = vbWindowBackground
            Else
                txtField(6).BackColor = vbButtonFace
                txtField(7).BackColor = vbButtonFace
            End If
            txtField(6).enabled = chkData(Index).Value = vbChecked
            txtField(7).enabled = chkData(Index).Value = vbChecked
            
        Case 4
            If chkData(Index).Value = vbChecked Then
                txtField(8).BackColor = vbWindowBackground
                txtField(9).BackColor = vbWindowBackground
            Else
                txtField(8).BackColor = vbButtonFace
                txtField(9).BackColor = vbButtonFace
            End If
            txtField(8).enabled = chkData(Index).Value = vbChecked
            txtField(9).enabled = chkData(Index).Value = vbChecked
    End Select
End Sub

Private Sub cmdAction_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Dim lRecordsAffected As Long, lTotalRecordsAffected As Long
    Dim iTrans As Integer
    
    Select Case Index
        Case 0
            iTrans = MainDB.BeginTrans
            
            lTotalRecordsAffected = 0
            
            Call ShowStatusBar("Menghapus data alumni...")
            If chkData(0).Value = vbChecked Then
                'Alumni
                sSQL = "DELETE FROM Alumni WHERE NoAlumni BETWEEN '" & txtField(0).Text & "' AND '" & txtField(1).Text & "'"
                MainDB.Execute sSQL, lRecordsAffected, Options:=adCmdText
            End If
            lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
                        
            Call ShowStatusBar("Menghapus data gereja dan organisasi kristen lainnya...")
            lRecordsAffected = 0
            If chkData(1).Value = vbChecked Then
                'Gereja dan Organisasi Kristen
                sSQL = "DELETE FROM Gereja WHERE NoArsip BETWEEN '" & txtField(2).Text & "' AND '" & txtField(3).Text & "'"
                MainDB.Execute sSQL, lRecordsAffected, Options:=adCmdText
            End If
            lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
            
            Call ShowStatusBar("Menghapus data pembicara...")
            lRecordsAffected = 0
            If chkData(2).Value = vbChecked Then
                'Pembicara
                sSQL = "DELETE FROM Pembicara WHERE NoArsip BETWEEN '" & txtField(4).Text & "' AND '" & txtField(5).Text & "'"
                MainDB.Execute sSQL, lRecordsAffected, Options:=adCmdText
            End If
            lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
            
            Call ShowStatusBar("Menghapus data tempat retret...")
            lRecordsAffected = 0
            If chkData(3).Value = vbChecked Then
                'Tempat Retret
                sSQL = "DELETE FROM TempatRetret WHERE NoArsip BETWEEN '" & txtField(6).Text & "' AND '" & txtField(7).Text & "'"
                MainDB.Execute sSQL, lRecordsAffected, Options:=adCmdText
            End If
            lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
            
            Call ShowStatusBar("Menghapus data PAKJ...")
            lRecordsAffected = 0
            If chkData(4).Value = vbChecked Then
                'PAKJ
                sSQL = "DELETE FROM PAKJ WHERE NoArsip BETWEEN '" & txtField(8).Text & "' AND '" & txtField(9).Text & "'"
                MainDB.Execute sSQL, lRecordsAffected, Options:=adCmdText
            End If
            lTotalRecordsAffected = lTotalRecordsAffected + lRecordsAffected
            
            If iTrans > 0 Then MainDB.CommitTrans
            Call RefreshDatabaseCache(MainDB)
            Call ShowStatusBar("RESET")
            
            MsgBox lTotalRecordsAffected & " record(s) affected.", vbInformation, Caption
        Case 1
            Unload Me
    End Select
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, Caption
    If iTrans > 0 Then MainDB.RollbackTrans
    Call ShowStatusBar("RESET")
End Sub

Private Sub cmdPrompt_Click(Index As Integer)
    Dim PromptTitle As String, ColumnsTitle As Variant
    Select Case Index
        Case 0, 1 'Alumni
            sSQL = "SELECT NoAlumni, NamaLengkap FROM Alumni ORDER BY NoAlumni"
            PromptTitle = "Prompt Alumni"
            ColumnsTitle = Array("No. Alumni", "Nama")
        Case 2, 3 'Gereja dan Organisasi Kristen
            sSQL = "SELECT NoArsip, NamaOrganisasi FROM Gereja ORDER BY NoArsip"
            PromptTitle = "Prompt Gereja & Organisasi Kristen"
            ColumnsTitle = Array("No. Arsip", "Nama")
        Case 4, 5 'Pembicara
            sSQL = "SELECT NoArsip, NamaLengkap FROM Pembicara ORDER BY NoArsip"
            PromptTitle = "Prompt Pembicara"
            ColumnsTitle = Array("No. Arsip", "Nama")
        Case 6, 7 'Tempat Retret
            sSQL = "SELECT NoArsip, NamaTempat FROM TempatRetret ORDER BY NoArsip"
            PromptTitle = "Prompt Tempat Retret"
            ColumnsTitle = Array("No. Arsip", "Nama")
        Case 8, 9 'PAKJ
            sSQL = "SELECT NoArsip, NamaPersekutuan FROM PAKJ ORDER BY NoArsip"
            PromptTitle = "Prompt PAKJ"
            ColumnsTitle = Array("No. Arsip", "Nama")
    End Select
    If oDialog.ShowPrompt(ConnectString, sSQL, PromptTitle, ColumnsTitle, txtField(Index).Text) = DialogAnswerOK Then
        txtField(Index).Text = oDialog.ColumnValue(0)
    End If
End Sub
