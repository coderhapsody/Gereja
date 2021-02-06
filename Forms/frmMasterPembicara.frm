VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMasterPembicara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembicara"
   ClientHeight    =   6555
   ClientLeft      =   1050
   ClientTop       =   1485
   ClientWidth     =   9525
   Icon            =   "frmMasterPembicara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9525
   Begin VB.Frame fraFrame 
      Height          =   6585
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   -45
      Width           =   9510
      Begin VB.CommandButton cmdButton 
         Caption         =   "Catatan"
         Height          =   825
         Index           =   1
         Left            =   1485
         Picture         =   "frmMasterPembicara.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   5355
         Width           =   2490
      End
      Begin VB.TextBox txtField 
         BackColor       =   &H8000000F&
         DataField       =   "LastUpdate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, dd mmmm yyyy   hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "dtcMaindata"
         Enabled         =   0   'False
         Height          =   330
         Index           =   11
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   4635
         Width           =   3480
      End
      Begin VB.TextBox txtField 
         DataField       =   "AlamatKantor"
         DataSource      =   "dtcMaindata"
         Height          =   690
         Index           =   10
         Left            =   1485
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3915
         Width           =   3480
      End
      Begin VB.TextBox txtField 
         DataField       =   "KodePos"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   9
         Left            =   1485
         TabIndex        =   9
         Top             =   3555
         Width           =   960
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kota"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   8
         Left            =   1485
         TabIndex        =   8
         Top             =   3195
         Width           =   1410
      End
      Begin VB.TextBox txtField 
         DataField       =   "RW"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   7
         Left            =   2655
         TabIndex        =   7
         Top             =   2835
         Width           =   600
      End
      Begin VB.TextBox txtField 
         DataField       =   "RT"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   6
         Left            =   1485
         TabIndex        =   6
         Top             =   2835
         Width           =   600
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kecamatan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   5
         Left            =   1485
         TabIndex        =   5
         Top             =   2475
         Width           =   1635
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kelurahan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   4
         Left            =   1485
         TabIndex        =   4
         Top             =   2115
         Width           =   1635
      End
      Begin VB.TextBox txtField 
         DataField       =   "AlamatRumah"
         DataSource      =   "dtcMaindata"
         Height          =   690
         Index           =   3
         Left            =   1485
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1395
         Width           =   3480
      End
      Begin VB.TextBox txtField 
         DataField       =   "NamaLengkap"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   1
         Left            =   1485
         TabIndex        =   1
         Top             =   675
         Width           =   3480
      End
      Begin VB.TextBox txtField 
         DataField       =   "NamaPanggilan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   2
         Left            =   1485
         TabIndex        =   2
         Top             =   1035
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   0
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox txtField 
         DataField       =   "NoArsip"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   0
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   1545
      End
      Begin VB.CommandButton cmdAutoNumber 
         Caption         =   "&Auto Number"
         Height          =   330
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   315
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   1
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4500
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Hapus Foto"
         Height          =   330
         Index           =   0
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   6030
         Width           =   1455
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Index           =   2
         Left            =   5265
         TabIndex        =   37
         Top             =   225
         Width           =   4020
         Begin VB.TextBox txtField 
            DataField       =   "Email1"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   12
            Left            =   1260
            TabIndex        =   11
            Top             =   360
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Email2"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   13
            Left            =   1260
            TabIndex        =   12
            Top             =   720
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Email3"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   14
            Left            =   1260
            TabIndex        =   13
            Top             =   1080
            Width           =   2580
         End
         Begin VB.Label lblField 
            Caption         =   "Label Email 2"
            Height          =   195
            Index           =   16
            Left            =   135
            TabIndex        =   40
            Top             =   765
            Width           =   1050
         End
         Begin VB.Label lblField 
            Caption         =   "Label Email 3"
            Height          =   195
            Index           =   17
            Left            =   135
            TabIndex        =   39
            Top             =   1125
            Width           =   1050
         End
         Begin VB.Label lblField 
            Caption         =   "Label Email 1"
            Height          =   240
            Index           =   18
            Left            =   135
            TabIndex        =   38
            Top             =   405
            Width           =   1050
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Index           =   3
         Left            =   5265
         TabIndex        =   31
         Top             =   1935
         Width           =   4020
         Begin VB.TextBox txtField 
            DataField       =   "Telepon3"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   15
            Left            =   1215
            TabIndex        =   16
            Top             =   1080
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon2"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   16
            Left            =   1215
            TabIndex        =   15
            Top             =   720
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon1"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   17
            Left            =   1215
            TabIndex        =   14
            Top             =   360
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon4"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   18
            Left            =   1215
            TabIndex        =   17
            Top             =   1440
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon5"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   19
            Left            =   1215
            TabIndex        =   18
            Top             =   1800
            Width           =   2580
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 1"
            Height          =   240
            Index           =   19
            Left            =   135
            TabIndex        =   36
            Top             =   405
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 3"
            Height          =   195
            Index           =   20
            Left            =   135
            TabIndex        =   35
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 2"
            Height          =   195
            Index           =   21
            Left            =   135
            TabIndex        =   34
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 4"
            Height          =   195
            Index           =   22
            Left            =   135
            TabIndex        =   33
            Top             =   1485
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 5"
            Height          =   195
            Index           =   23
            Left            =   135
            TabIndex        =   32
            Top             =   1845
            Width           =   1005
         End
      End
      Begin VB.Label lblField 
         Caption         =   "Perubahan terakhir"
         Height          =   420
         Index           =   11
         Left            =   180
         TabIndex        =   47
         Top             =   4590
         Width           =   1095
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   6570
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Foto"
         Height          =   195
         Index           =   24
         Left            =   5985
         TabIndex        =   43
         Top             =   4545
         Width           =   465
      End
      Begin VB.Label lblField 
         Caption         =   "Alamat Kantor"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   30
         Top             =   3915
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "Kode Pos"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   29
         Top             =   3555
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "Kota"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   28
         Top             =   3195
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "RW"
         Height          =   195
         Index           =   7
         Left            =   2295
         TabIndex        =   27
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblField 
         Caption         =   "RT"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   26
         Top             =   2835
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "Kecamatan"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "Kelurahan"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "Alamat Rumah"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   23
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Panggilan"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Lengkap"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblField 
         Caption         =   "No. Arsip"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc dtcMaindata 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6225
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
End
Attribute VB_Name = "frmMasterPembicara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ITransaction

Private bInternalOperation As Boolean
Private bDataChanged As Boolean
Private bIsAutoNumberFetched As Boolean

Private sFotoFileName As String
Private oAutoNum As CAutoNumber

Private Sub cmdAutoNumber_Click()
    If DataChanged Then
        If Trim$(txtField(0).Text) = BLANK Then
            txtField(0).Text = oAutoNum.FetchAutoNumber(MASTER_PEMBICARA, BLANK, dtcMaindata.Recordset("NoArsip").DefinedSize - Len(Parameter![PrefixArsipPembicara]))
            Call oAutoNum.IncrementAutoNumber(MASTER_PEMBICARA)
            bIsAutoNumberFetched = True
        End If
    End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
    '// Beginning of template
    If IsBOFEOF(dtcMaindata) Then Exit Sub
    '//
    
    'Your code here...
    
    Select Case Index
        Case 0
            Set imgFoto.Picture = LoadPicture()
            sFotoFileName = BLANK
            dtcMaindata.Recordset![Foto] = BLANK
            DataChanged = True
        Case 1
            bDataChanged = DataChanged
            Load frmCatatanPembicara
            frmCatatanPembicara.NoArsip = txtField(0).Text
            Call frmCatatanPembicara.dtcMaindata_Refresh
            frmCatatanPembicara.Show vbModal
            DataChanged = bDataChanged
    End Select
End Sub

Private Sub cmdPrompt_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Select Case Index
        Case 0
            If DataChanged Then Exit Sub
            sSQL = "SELECT NoArsip, NamaLengkap FROM Pembicara ORDER BY NoArsip"
            If oDialog.ShowPrompt(ConnectString, sSQL, "Prompt No. Arsip Pembicara", _
                Array("No. Arsip", "Nama"), txtField(Index).Text) = DialogAnswerOK Then
                Call dtcMaindata_Refresh(oDialog.ColumnValue(0))
            End If
        Case 1
            If IsBOFEOF(dtcMaindata) Then Exit Sub
            With frmMain.dlgMain
                .Filter = "Pictures|*.bmp;*.jpg;*.gif;*.png"
                .CancelError = False
                .DialogTitle = "Pilih file foto yang akan digunakan..."
                .ShowOpen
                If Trim$(.FileName) <> BLANK Then
                    dtcMaindata.Recordset![Foto] = .FileName
                    sFotoFileName = .FileName
                    Set imgFoto.Picture = LoadPicture(.FileName)
                    DataChanged = True
                End If
            End With
    End Select
    
    Exit Sub
        
ErrHandler:
    MsgBox Err.Description, vbExclamation, Caption
End Sub

Private Sub dtcMaindata_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    
    bInternalOperation = True
    If Not IsBOFEOF(pRecordset) Then
        If dtcMaindata.Recordset.EditMode = adEditNone Then
            
            
            If Trim$(dtcMaindata.Recordset![Foto] & "") <> BLANK Then
                Set imgFoto.Picture = LoadPicture(dtcMaindata.Recordset![Foto])
            Else
                Set imgFoto.Picture = LoadPicture()
            End If
        End If
    Else
        Set imgFoto.Picture = LoadPicture()
    End If
    bInternalOperation = False

End Sub

Private Sub Form_Load()
    '// Beginning of Template
    Call LoadPosition(Me, INIPath)
    Call dtcMaindata_Refresh(BLANK)
    '//
    
    
    'Your code here...
    
    Set oAutoNum = New CAutoNumber
    Set oAutoNum.DatabaseConnection = MainDB
    
    lblField(18).Caption = Trim$(Parameter![LabelEmail1] & "")
    If lblField(18).Caption = BLANK Then
        lblField(18).Visible = False
        txtField(12).Visible = False
    End If
    
    
    lblField(16).Caption = Trim$(Parameter![LabelEmail2] & "")
    If lblField(16).Caption = BLANK Then
        lblField(16).Visible = False
        txtField(13).Visible = False
    End If
    
    lblField(17).Caption = Trim$(Parameter![LabelEmail3] & "")
    If lblField(17).Caption = BLANK Then
        lblField(16).Visible = False
        txtField(14).Visible = False
    End If
    
    lblField(19).Caption = Trim$(Parameter![LabelTelepon1] & "")
    If lblField(19).Caption = BLANK Then
        lblField(19).Visible = False
        txtField(17).Visible = False
    End If
    
    lblField(21).Caption = Trim$(Parameter![LabelTelepon2] & "")
    If lblField(21).Caption = BLANK Then
        lblField(21).Visible = False
        txtField(16).Visible = False
    End If
    
    lblField(20).Caption = Trim$(Parameter![LabelTelepon3] & "")
    If lblField(20).Caption = BLANK Then
        lblField(20).Visible = False
        txtField(15).Visible = False
    End If
    
    lblField(22).Caption = Trim$(Parameter![LabelTelepon4] & "")
    If lblField(22).Caption = BLANK Then
        lblField(22).Visible = False
        txtField(18).Visible = False
    End If
    
    lblField(23).Caption = Trim$(Parameter![LabelTelepon5] & "")
    If lblField(23).Caption = BLANK Then
        lblField(23).Visible = False
        txtField(19).Visible = False
    End If
    
    For iLoop = cmdPrompt.LBound To cmdPrompt.UBound
        Set cmdPrompt(iLoop).Picture = oPromptIcon
    Next
    
    If Parameter![AutoNumber] Then
        cmdAutoNumber.enabled = True
        txtField(0).Locked = True
    Else
        cmdAutoNumber.enabled = False
        txtField(0).Locked = False
    End If
    
    '// Beginning of template
    DataChanged = False
    '//
End Sub

Private Sub Form_Activate()
    '// Beginning of template
    DataChanged = bDataChanged
    '//
    
    'Your code here..
End Sub

Private Sub Form_Deactivate()

    'Your code here..
    
    '// Beginning of template
    bDataChanged = DataChanged
    '//
End Sub

Private Sub Form_GotFocus()
    '// Beginning of template
    DataChanged = bDataChanged
    '//
    
    'Your code here...
End Sub

Private Sub Form_LostFocus()
    'Your code here..
    
    '// Beginning of template
    bDataChanged = DataChanged
    '//
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '// Beginning of template
    Call SavePosition(Me, INIPath)
    Call DisableToolbarButtons
    '//
    
    'Do resource clean up here...
    Set oAutoNum = Nothing
End Sub

Public Sub ITransaction_MasterAddNew()
    Dim Control As Control
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        ElseIf TypeOf Control Is ComboBox Then
            Control.ListIndex = -1
        ElseIf TypeOf Control Is MaskEdBox Then
            Control.PromptInclude = False
            Control.Text = ""
            Control.PromptInclude = True
        ElseIf TypeOf Control Is Image Then
            Set Control.Picture = Nothing
        End If
    Next

    bIsAutoNumberFetched = False
    dtcMaindata.Recordset.AddNew
    txtField(0).SetFocus
    DataChanged = True
End Sub

Public Sub ITransaction_MasterCancel()
    dtcMaindata.Recordset.Requery
    If bIsAutoNumberFetched Then
        Call oAutoNum.DecrementAutoNumber(MASTER_PEMBICARA)
    End If
    bIsAutoNumberFetched = False
    DataChanged = False
End Sub

Public Sub ITransaction_MasterDelete()
    If IsBOFEOF(dtcMaindata) Then Exit Sub
    
    If MsgBox("Hapus record ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    dtcMaindata.Recordset.Delete
    dtcMaindata.Recordset.UpdateBatch
    dtcMaindata.Recordset.Requery
End Sub

Public Sub ITransaction_MasterPrint()
    frmLapDaftarPembicara.Show
    frmLapDaftarPembicara.ZOrder ZOrderConstants.vbBringToFront
End Sub

Public Sub ITransaction_MasterRefresh()
    Call ShowStatusBar("Reloading record from database...")
    Call dtcMaindata_Refresh(txtField(0).Text)
    Call ShowStatusBar("RESET")
End Sub

Public Sub ITransaction_MasterSave()
    Dim lRecordsAffected As Long
    
    On Error GoTo ErrHandler
    
    If Not DataIsValid Then
        Exit Sub
    End If
        
    dtcMaindata.Recordset.Move 0
    dtcMaindata.Recordset.UpdateBatch
    
    MainDB.BeginTrans
    sSQL = "UPDATE Pembicara SET LastUpdate=Now WHERE NoArsip = '" & txtField(0).Text & "'"
    MainDB.Execute sSQL, lRecordsAffected, adCmdText
    MainDB.CommitTrans
    Call RefreshDatabaseCache(MainDB)
        
    bIsAutoNumberFetched = False
    DataChanged = False
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, Caption
End Sub

Private Property Let DataChanged(ByVal Toggle As Boolean)
    '// Beginning of template
    Call AdjustToolbarButton(Not Toggle, Toggle, Not Toggle, Toggle)
    '//
End Property

Private Property Get DataChanged() As Boolean
    '// Beginning of template
    DataChanged = frmMain.Toolbar.Buttons("SAVE").enabled
    '//
End Property

Private Function DataIsValid() As Boolean
    DataIsValid = True
End Function

Public Sub dtcMaindata_Refresh(ByVal NoArsip As String)
    With dtcMaindata
        .ConnectionString = ConnectString
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeout
        .LockType = adLockBatchOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .RecordSource = "SELECT * FROM Pembicara WHERE NoArsip = '" & NoArsip & "'"
        .Refresh
    End With
    
    '// Beginning of Template
    Dim Control As Control
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.DataField <> BLANK Then
                If dtcMaindata.Recordset.Fields(Control.DataField).Type <> adDate Then
                        Control.MaxLength = dtcMaindata.Recordset.Fields(Control.DataField).DefinedSize
                End If
            End If
        End If
    Next
    '//
End Sub

Private Sub txtField_Change(Index As Integer)
    '// Beginning of template
    If txtField(Index).DataChanged Then DataChanged = True
    '//
    
    'Your code here...
End Sub

Private Sub txtField_GotFocus(Index As Integer)
    '// Beginning of template
    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index).Text)
    '//
    
    'Your code here...
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
    '// Beginning of template
    If IsBOFEOF(dtcMaindata) Then KeyAscii = 0
    KeyAscii = CheckKeyPress(KeyAscii)
    '//
    
    'Your code here...
End Sub
