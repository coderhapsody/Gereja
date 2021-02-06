VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmKonfigurasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi"
   ClientHeight    =   5565
   ClientLeft      =   825
   ClientTop       =   2655
   ClientWidth     =   9585
   Icon            =   "frmKonfigurasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9585
   Begin TabDlg.SSTab Tab 
      Height          =   5415
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Umum"
      TabPicture(0)   =   "frmKonfigurasi.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkAutoNumber"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFrame(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFrame(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkSideBar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Pendataan"
      TabPicture(1)   =   "frmKonfigurasi.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraFrame(6)"
      Tab(1).Control(1)=   "fraFrame(5)"
      Tab(1).Control(2)=   "fraFrame(4)"
      Tab(1).Control(3)=   "fraFrame(3)"
      Tab(1).Control(4)=   "fraFrame(2)"
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chkSideBar 
         Caption         =   "Tampilkan Shortcut"
         Height          =   195
         Left            =   405
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2580
      End
      Begin VB.Frame fraFrame 
         Caption         =   "PAKJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   6
         Left            =   -74775
         TabIndex        =   42
         Top             =   3735
         Width           =   4200
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            DataField       =   "AutoNumberPAKJ"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   17
            Left            =   2700
            TabIndex        =   13
            Top             =   810
            Width           =   1230
         End
         Begin VB.TextBox txtField 
            DataField       =   "PrefixArsipPAKJ"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   16
            Left            =   2700
            TabIndex        =   12
            Top             =   360
            Width           =   1230
         End
         Begin VB.Image imgImage 
            Height          =   1005
            Index           =   4
            Left            =   180
            MouseIcon       =   "frmKonfigurasi.frx":047A
            MousePointer    =   99  'Custom
            Picture         =   "frmKonfigurasi.frx":0784
            Stretch         =   -1  'True
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Nomor Terakhir"
            Height          =   195
            Index           =   18
            Left            =   1350
            TabIndex        =   44
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label lblField 
            Caption         =   "Prefix No."
            Height          =   195
            Index           =   17
            Left            =   1350
            TabIndex        =   43
            Top             =   405
            Width           =   1320
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Tempat Retret"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   5
         Left            =   -70275
         TabIndex        =   39
         Top             =   2160
         Width           =   4515
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            DataField       =   "AutoNumberTempatRetret"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   15
            Left            =   3150
            TabIndex        =   17
            Top             =   810
            Width           =   1230
         End
         Begin VB.TextBox txtField 
            DataField       =   "PrefixArsipTempatRetret"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   14
            Left            =   3150
            TabIndex        =   16
            Top             =   360
            Width           =   1230
         End
         Begin VB.Image imgImage 
            Height          =   1005
            Index           =   3
            Left            =   315
            MouseIcon       =   "frmKonfigurasi.frx":160C
            MousePointer    =   99  'Custom
            Picture         =   "frmKonfigurasi.frx":1916
            Stretch         =   -1  'True
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Nomor Terakhir"
            Height          =   195
            Index           =   16
            Left            =   1800
            TabIndex        =   41
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label lblField 
            Caption         =   "Prefix"
            Height          =   195
            Index           =   15
            Left            =   1800
            TabIndex        =   40
            Top             =   405
            Width           =   1320
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Pembicara"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   4
         Left            =   -70275
         TabIndex        =   36
         Top             =   630
         Width           =   4515
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            DataField       =   "AutoNumberPembicara"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   13
            Left            =   3150
            TabIndex        =   15
            Top             =   810
            Width           =   1230
         End
         Begin VB.TextBox txtField 
            DataField       =   "PrefixArsipPembicara"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   12
            Left            =   3150
            TabIndex        =   14
            Top             =   360
            Width           =   1230
         End
         Begin VB.Image imgImage 
            Height          =   1005
            Index           =   2
            Left            =   270
            MouseIcon       =   "frmKonfigurasi.frx":61E8
            MousePointer    =   99  'Custom
            Picture         =   "frmKonfigurasi.frx":64F2
            Stretch         =   -1  'True
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Nomor Terakhir"
            Height          =   195
            Index           =   14
            Left            =   1800
            TabIndex        =   38
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label lblField 
            Caption         =   "Prefix"
            Height          =   195
            Index           =   13
            Left            =   1800
            TabIndex        =   37
            Top             =   405
            Width           =   1320
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Gereja dan Organisasi Kristen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   3
         Left            =   -74775
         TabIndex        =   33
         Top             =   2160
         Width           =   4200
         Begin VB.TextBox txtField 
            DataField       =   "PrefixArsipOrganisasi"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   11
            Left            =   2700
            TabIndex        =   10
            Top             =   450
            Width           =   1230
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            DataField       =   "AutoNumberOrganisasi"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   10
            Left            =   2700
            TabIndex        =   11
            Top             =   855
            Width           =   1230
         End
         Begin VB.Image imgImage 
            Height          =   1005
            Index           =   1
            Left            =   180
            MouseIcon       =   "frmKonfigurasi.frx":A169
            MousePointer    =   99  'Custom
            Picture         =   "frmKonfigurasi.frx":A473
            Stretch         =   -1  'True
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Prefix"
            Height          =   195
            Index           =   12
            Left            =   1440
            TabIndex        =   35
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lblField 
            Caption         =   "Nomor Terakhir"
            Height          =   195
            Index           =   11
            Left            =   1440
            TabIndex        =   34
            Top             =   855
            Width           =   1185
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Alumni"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Index           =   2
         Left            =   -74730
         TabIndex        =   30
         Top             =   630
         Width           =   4155
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            DataField       =   "CurrentAutoNumber"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   9
            Left            =   2700
            TabIndex        =   9
            Top             =   855
            Width           =   1230
         End
         Begin VB.TextBox txtField 
            DataField       =   "PrefixNoAlumni"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   0
            Left            =   2700
            TabIndex        =   8
            Top             =   405
            Width           =   1230
         End
         Begin VB.Image imgImage 
            Height          =   1005
            Index           =   0
            Left            =   180
            MouseIcon       =   "frmKonfigurasi.frx":E23E
            MousePointer    =   99  'Custom
            Picture         =   "frmKonfigurasi.frx":E548
            Stretch         =   -1  'True
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Nomor Terakhir"
            Height          =   195
            Index           =   10
            Left            =   1350
            TabIndex        =   32
            Top             =   855
            Width           =   1185
         End
         Begin VB.Label lblField 
            Caption         =   "Prefix"
            Height          =   195
            Index           =   0
            Left            =   1350
            TabIndex        =   31
            Top             =   450
            Width           =   1320
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Label Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Index           =   0
         Left            =   4725
         TabIndex        =   26
         Top             =   1530
         Width           =   4200
         Begin VB.TextBox txtField 
            DataField       =   "LabelEmail1"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   1
            Left            =   1080
            TabIndex        =   5
            Top             =   450
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon2"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   2
            Left            =   1080
            TabIndex        =   6
            Top             =   855
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon3"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   3
            Left            =   1080
            TabIndex        =   7
            Top             =   1260
            Width           =   2805
         End
         Begin VB.Label lblField 
            Caption         =   "Email 1"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   29
            Top             =   495
            Width           =   645
         End
         Begin VB.Label lblField 
            Caption         =   "Email 2"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   28
            Top             =   900
            Width           =   645
         End
         Begin VB.Label lblField 
            Caption         =   "Email 3"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   27
            Top             =   1305
            Width           =   645
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Label Telepon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Index           =   1
         Left            =   270
         TabIndex        =   20
         Top             =   1530
         Width           =   4200
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon5"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   8
            Left            =   1080
            TabIndex        =   4
            Top             =   2070
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon4"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   7
            Left            =   1080
            TabIndex        =   3
            Top             =   1665
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon1"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   6
            Left            =   1080
            TabIndex        =   0
            Top             =   450
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon2"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   5
            Left            =   1080
            TabIndex        =   1
            Top             =   855
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            DataField       =   "LabelTelepon3"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   4
            Left            =   1080
            TabIndex        =   2
            Top             =   1260
            Width           =   2805
         End
         Begin VB.Label lblField 
            Caption         =   "Telepon 5"
            Height          =   195
            Index           =   8
            Left            =   225
            TabIndex        =   25
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblField 
            Caption         =   "Telepon 4"
            Height          =   195
            Index           =   7
            Left            =   225
            TabIndex        =   24
            Top             =   1710
            Width           =   735
         End
         Begin VB.Label lblField 
            Caption         =   "Telepon 3"
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   23
            Top             =   1305
            Width           =   735
         End
         Begin VB.Label lblField 
            Caption         =   "Telepon 2"
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   22
            Top             =   900
            Width           =   735
         End
         Begin VB.Label lblField 
            Caption         =   "Telepon 1"
            Height          =   195
            Index           =   6
            Left            =   225
            TabIndex        =   21
            Top             =   495
            Width           =   735
         End
      End
      Begin VB.CheckBox chkAutoNumber 
         Caption         =   "Penomoran &Otomatis"
         Height          =   195
         Left            =   405
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   675
         Width           =   2580
      End
   End
   Begin MSAdodcLib.Adodc dtcMaindata 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5235
      Visible         =   0   'False
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
Attribute VB_Name = "frmKonfigurasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ITransaction

Private bInternalOperation As Boolean
Private bDataChanged As Boolean

Private Sub chkAutoNumber_Click()
    DataChanged = True
End Sub

Private Sub chkSideBar_Click()
    DataChanged = True
End Sub

Private Sub dtcMaindata_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    
    bInternalOperation = True
    If Not IsBOFEOF(pRecordset) Then
        If dtcMaindata.Recordset.EditMode = adEditNone Then
            If pRecordset![SideBar] = True Then
                chkSideBar.Value = vbChecked
            Else
                chkSideBar.Value = vbUnchecked
            End If
            
            If pRecordset![AutoNumber] = True Then
                chkAutoNumber.Value = vbChecked
            Else
                chkAutoNumber.Value = vbUnchecked
            End If
        End If
    Else
    End If
    bInternalOperation = False
End Sub

Private Sub Form_Load()
    
    '// Beginning of Template
    Call LoadPosition(Me, INIPath)
    Call dtcMaindata_Refresh
    '//
    
    
    'Your code here...
    For iLoop = frmMain.mnuMain.LBound To frmMain.mnuMain.UBound
        frmMain.mnuMain(iLoop).enabled = False
        frmMain.picSideBar.enabled = False
    Next
    
    
    '// Beginning of template
    DataChanged = False
    '//
End Sub

Public Sub dtcMaindata_Refresh()
    With dtcMaindata
        .ConnectionString = ConnectString
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeout
        .LockType = adLockBatchOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .RecordSource = "SELECT * FROM Parameter"
        .Refresh
    End With
        
    '// Beginning of Template
    Dim Control As Control
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.DataField <> BLANK Then
                Control.MaxLength = dtcMaindata.Recordset.Fields(Control.DataField).DefinedSize
            End If
        End If
    Next
    '//
End Sub

Private Property Get DataIsValid() As Boolean
    DataIsValid = True
End Property

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
    Call RefreshDatabaseCache(MainDB)
    Parameter.Requery
    frmMain.picSideBar.Visible = Parameter![SideBar]
    
    For iLoop = frmMain.mnuMain.LBound To frmMain.mnuMain.UBound
        frmMain.mnuMain(iLoop).enabled = True
        frmMain.picSideBar.enabled = True
    Next
    
    DoEvents
End Sub

Public Sub ITransaction_MasterAddNew()
End Sub

Public Sub ITransaction_MasterCancel()
    dtcMaindata.Recordset.Requery
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
End Sub

Public Sub ITransaction_MasterRefresh()
    Call ShowStatusBar("Reloading record from database...")
    Call dtcMaindata_Refresh
    Call ShowStatusBar("RESET")
End Sub

Public Sub ITransaction_MasterSave()
    Dim lRecordsAffected As Long
    On Error GoTo ErrHandler
    
    If Not DataIsValid Then
        Exit Sub
    End If
    
    If chkAutoNumber.Value = vbChecked Then dtcMaindata.Recordset![AutoNumber] = True Else dtcMaindata.Recordset![AutoNumber] = False
    If chkSideBar.Value = vbChecked Then dtcMaindata.Recordset![SideBar] = True Else dtcMaindata.Recordset![SideBar] = False
          
    dtcMaindata.Recordset.UpdateBatch

    Call RefreshDatabaseCache(MainDB)
    Parameter.Requery
    frmMain.picSideBar.Visible = chkSideBar.Value = vbChecked
    
    DataChanged = False
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, Caption
    Resume
End Sub

Private Property Let DataChanged(ByVal Toggle As Boolean)
    '// Beginning of template
    Call AdjustToolbarButton(False, Toggle, False, Toggle, False, True)
    '//
End Property

Private Property Get DataChanged() As Boolean
    '// Beginning of template
    DataChanged = frmMain.Toolbar.Buttons("SAVE").enabled
    '//
End Property

Private Sub mskField_Change()
    '// Beginning of template
    If bInternalOperation Then Exit Sub
    If Not DataChanged Then DataChanged = True
    '//
    
    'Your code here...
End Sub

Private Sub mskField_KeyPress(KeyAscii As Integer)
    '// Beginning of template
    If IsBOFEOF(dtcMaindata) Then KeyAscii = 0
    KeyAscii = CheckKeyPress(KeyAscii)
    '//
    
    'Your code here
    KeyAscii = NumberOnly(KeyAscii)
End Sub

Private Sub optJenisKelamin_Click(Index As Integer)
    '// Beginning of template
    If bInternalOperation Then Exit Sub
    '//
    
    'Your code here...
    
    '// Beginning of template
    DataChanged = True
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


