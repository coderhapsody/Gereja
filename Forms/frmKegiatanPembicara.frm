VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCatatanPembicara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catatan"
   ClientHeight    =   5775
   ClientLeft      =   2775
   ClientTop       =   2385
   ClientWidth     =   8775
   Icon            =   "frmKegiatanPembicara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8775
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Batal"
      Height          =   375
      Index           =   1
      Left            =   7065
      TabIndex        =   2
      Top             =   5265
      Width           =   1635
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   7065
      TabIndex        =   1
      Top             =   4815
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc dtcMaindata 
      Height          =   330
      Left            =   7155
      Top             =   4320
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtField 
      DataField       =   "Catatan"
      DataSource      =   "dtcMaindata"
      Height          =   5460
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   6810
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   7155
      Picture         =   "frmKegiatanPembicara.frx":0CCA
      Stretch         =   -1  'True
      Top             =   225
      Width           =   1305
   End
End
Attribute VB_Name = "frmCatatanPembicara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_NoArsip As String

Public Property Let NoArsip(ByVal NoArsip As String)
    m_NoArsip = NoArsip
End Property

Public Property Get NoArsip() As String
    NoArsip = m_NoArsip
End Property

Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0 'Simpan
            dtcMaindata.Recordset.UpdateBatch
            Unload Me
        Case 1 'Batal
            If MsgBox("Batal semua perubahan pada catatan ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Unload Me
                frmMasterPembicara.ZOrder ZOrderConstants.vbBringToFront
            End If
    End Select
End Sub

Private Sub Form_Load()
    Call DisableToolbarButtons
End Sub

Public Sub dtcMaindata_Refresh()
    With dtcMaindata
        .ConnectionString = ConnectString
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeout
        .LockType = adLockBatchOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .RecordSource = "SELECT * FROM Pembicara WHERE NoArsip='" & NoArsip & "'"
        .Refresh
    End With
End Sub

Private Sub txtField_Change()
    cmdButton(0).enabled = txtField.DataChanged
End Sub
