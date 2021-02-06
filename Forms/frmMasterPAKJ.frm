VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMasterPAKJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Persekutuan Alumni Kristen Jakarta (PAKJ)"
   ClientHeight    =   6840
   ClientLeft      =   1410
   ClientTop       =   2835
   ClientWidth     =   9990
   Icon            =   "frmMasterPAKJ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9990
   Begin VB.Frame fraFrame 
      Height          =   6900
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   -45
      Width           =   10005
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   1
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   4860
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Hapus Foto"
         Height          =   330
         Index           =   0
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   6390
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   2
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4860
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Hapus Foto"
         Height          =   330
         Index           =   1
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6390
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   3
         Left            =   8910
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   4860
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Hapus Foto"
         Height          =   330
         Index           =   2
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   6390
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         DataField       =   "AlamatSekretariat"
         DataSource      =   "dtcMaindata"
         Height          =   510
         Index           =   2
         Left            =   1665
         TabIndex        =   2
         Top             =   1035
         Width           =   3030
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Index           =   1
         Left            =   5130
         TabIndex        =   25
         Top             =   225
         Width           =   4650
         Begin VB.TextBox txtField 
            DataField       =   "NamaCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   12
            Left            =   1170
            TabIndex        =   11
            Top             =   315
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "AlamatCP"
            DataSource      =   "dtcMaindata"
            Height          =   510
            Index           =   13
            Left            =   1170
            TabIndex        =   12
            Top             =   675
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "KelurahanCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   14
            Left            =   1170
            TabIndex        =   13
            Top             =   1215
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            DataField       =   "KecamatanCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   15
            Left            =   1170
            TabIndex        =   14
            Top             =   1575
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            DataField       =   "KotaCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   16
            Left            =   1170
            TabIndex        =   15
            Top             =   1935
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            DataField       =   "KodePosCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   17
            Left            =   1170
            TabIndex        =   16
            Top             =   2295
            Width           =   825
         End
         Begin VB.TextBox txtField 
            DataField       =   "EmailCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   18
            Left            =   1170
            TabIndex        =   17
            Top             =   2655
            Width           =   2400
         End
         Begin VB.TextBox txtField 
            DataField       =   "TeleponCP"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   19
            Left            =   1170
            TabIndex        =   18
            Top             =   3015
            Width           =   1545
         End
         Begin VB.Label lblField 
            Caption         =   "Nama"
            Height          =   195
            Index           =   14
            Left            =   225
            TabIndex        =   33
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblField 
            Caption         =   "Alamat"
            Height          =   195
            Index           =   15
            Left            =   225
            TabIndex        =   32
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label lblField 
            Caption         =   "Kecamatan"
            Height          =   195
            Index           =   16
            Left            =   225
            TabIndex        =   31
            Top             =   1575
            Width           =   1365
         End
         Begin VB.Label lblField 
            Caption         =   "Kelurahan"
            Height          =   195
            Index           =   17
            Left            =   225
            TabIndex        =   30
            Top             =   1215
            Width           =   1365
         End
         Begin VB.Label lblField 
            Caption         =   "Kode Pos"
            Height          =   195
            Index           =   18
            Left            =   225
            TabIndex        =   29
            Top             =   2340
            Width           =   1365
         End
         Begin VB.Label lblField 
            Caption         =   "Kota"
            Height          =   195
            Index           =   20
            Left            =   225
            TabIndex        =   28
            Top             =   1935
            Width           =   1365
         End
         Begin VB.Label lblField 
            Caption         =   "Telepon"
            Height          =   195
            Index           =   21
            Left            =   225
            TabIndex        =   27
            Top             =   3060
            Width           =   1365
         End
         Begin VB.Label lblField 
            Caption         =   "e-Mail"
            Height          =   195
            Index           =   22
            Left            =   225
            TabIndex        =   26
            Top             =   2700
            Width           =   1365
         End
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kelurahan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   3
         Left            =   1665
         TabIndex        =   3
         Top             =   1575
         Width           =   1545
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kecamatan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   4
         Left            =   1665
         TabIndex        =   4
         Top             =   1935
         Width           =   1545
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kota"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   5
         Left            =   1665
         TabIndex        =   5
         Top             =   2295
         Width           =   1545
      End
      Begin VB.TextBox txtField 
         DataField       =   "KodePos"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   7
         Left            =   1665
         TabIndex        =   6
         Top             =   2655
         Width           =   825
      End
      Begin VB.TextBox txtField 
         DataField       =   "Email"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   8
         Left            =   1665
         TabIndex        =   7
         Top             =   3015
         Width           =   2400
      End
      Begin VB.TextBox txtField 
         DataField       =   "Milis"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   9
         Left            =   1665
         TabIndex        =   8
         Top             =   3375
         Width           =   2400
      End
      Begin VB.TextBox txtField 
         DataField       =   "Telepon"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   10
         Left            =   1665
         TabIndex        =   9
         Top             =   3735
         Width           =   1545
      End
      Begin VB.TextBox txtField 
         DataField       =   "Fax"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   11
         Left            =   1665
         TabIndex        =   10
         Top             =   4095
         Width           =   1545
      End
      Begin VB.TextBox txtField 
         DataField       =   "NamaPersekutuan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   1
         Left            =   1665
         TabIndex        =   1
         Top             =   675
         Width           =   3030
      End
      Begin VB.CommandButton cmdAutoNumber 
         Caption         =   "&Auto Number"
         Height          =   330
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   315
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   0
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox txtField 
         DataField       =   "NoArsip"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   0
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   1545
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   0
         Left            =   1215
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Foto 1"
         Height          =   195
         Index           =   24
         Left            =   540
         TabIndex        =   50
         Top             =   4950
         Width           =   465
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   1
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Foto 2"
         Height          =   195
         Index           =   12
         Left            =   3735
         TabIndex        =   49
         Top             =   4950
         Width           =   465
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   2
         Left            =   7425
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Foto 3"
         Height          =   195
         Index           =   13
         Left            =   6840
         TabIndex        =   48
         Top             =   4950
         Width           =   465
      End
      Begin VB.Label lblField 
         Caption         =   "Kelurahan"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   41
         Top             =   1620
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Kecamatan"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   40
         Top             =   1980
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Kota"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   39
         Top             =   2340
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Kode Pos"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   38
         Top             =   2700
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "e-Mail"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   37
         Top             =   3060
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Milis"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   36
         Top             =   3420
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Telepon"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   35
         Top             =   3780
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Fax."
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   34
         Top             =   4140
         Width           =   1365
      End
      Begin VB.Label lblField 
         Caption         =   "Alamat Sekretariat"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   1125
         Width           =   1410
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Persekutuan"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   23
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label lblField 
         Caption         =   "No. Arsip"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   360
         Width           =   870
      End
   End
   Begin MSAdodcLib.Adodc dtcMaindata 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6510
      Width           =   9990
      _ExtentX        =   17621
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
      Caption         =   ""
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
Attribute VB_Name = "frmMasterPAKJ"
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
            txtField(0).Text = oAutoNum.FetchAutoNumber(MASTER_PAKJ, BLANK, dtcMaindata.Recordset("NoArsip").DefinedSize - Len(Parameter![PrefixArsipPAKJ]))
            Call oAutoNum.IncrementAutoNumber(MASTER_PAKJ)
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
        Case 0 To 2
            Set imgFoto(Index).Picture = LoadPicture()
            sFotoFileName = BLANK
            dtcMaindata.Recordset.Fields("Foto" & (Index + 1)).Value = BLANK
            DataChanged = True
    End Select
End Sub

Private Sub cmdPrompt_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Select Case Index
        Case 0
            If DataChanged Then Exit Sub
            sSQL = "SELECT NoArsip, NamaPersekutuan FROM PAKJ ORDER BY NoArsip"
            If oDialog.ShowPrompt(ConnectString, sSQL, "Prompt Arsip Persekutuan Alumni Jakarta", _
                Array("No. Arsip", "Nama Persekutuan"), txtField(Index).Text) = DialogAnswerOK Then
                Call dtcMaindata_Refresh(oDialog.ColumnValue(0))
            End If
        Case 1 To 3
            If IsBOFEOF(dtcMaindata) Then Exit Sub
            With frmMain.dlgMain
                .Filter = "Pictures|*.bmp;*.jpg;*.gif;*.png"
                .CancelError = False
                .DialogTitle = "Pilih file foto yang akan digunakan..."
                .ShowOpen
                If Trim$(.FileName) <> BLANK Then
                    dtcMaindata.Recordset.Fields("Foto" & Index).Value = .FileName
                    sFotoFileName = .FileName
                    Set imgFoto(Index - 1).Picture = LoadPicture(.FileName)
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
            
            
            If Trim$(dtcMaindata.Recordset![Foto1] & "") <> BLANK Then
                Set imgFoto(0).Picture = LoadPicture(dtcMaindata.Recordset![Foto1])
            Else
                Set imgFoto(0).Picture = LoadPicture()
            End If
            
            If Trim$(dtcMaindata.Recordset![Foto2] & "") <> BLANK Then
                Set imgFoto(1).Picture = LoadPicture(dtcMaindata.Recordset![Foto2])
            Else
                Set imgFoto(1).Picture = LoadPicture()
            End If
            
            If Trim$(dtcMaindata.Recordset![Foto3] & "") <> BLANK Then
                Set imgFoto(2).Picture = LoadPicture(dtcMaindata.Recordset![Foto3])
            Else
                Set imgFoto(2).Picture = LoadPicture()
            End If
        End If
    Else
        For iLoop = imgFoto.LBound To imgFoto.UBound
            Set imgFoto(iLoop).Picture = LoadPicture()
        Next
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
        Call oAutoNum.DecrementAutoNumber(MASTER_PAKJ)
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
    frmLapDaftarPAKJ.Show
    frmLapDaftarPAKJ.ZOrder ZOrderConstants.vbBringToFront
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
    sSQL = "UPDATE PAKJ SET LastUpdate=Now WHERE NoArsip = '" & txtField(0).Text & "'"
    MainDB.Execute sSQL, lRecordsAffected, adCmdText
    MainDB.CommitTrans
    Call RefreshDatabaseCache(MainDB)
        
    bIsAutoNumberFetched = False
    DataChanged = False
    
    Exit Sub

ErrHandler:
    
    If Err.Number = -2147467259 And Err.Source = "Microsoft JET Database Engine" Then
        Resume
    End If
    
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
        .RecordSource = "SELECT * FROM PAKJ WHERE NoArsip = '" & NoArsip & "'"
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


