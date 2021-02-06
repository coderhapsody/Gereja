VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMasterAlumni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alumni"
   ClientHeight    =   7185
   ClientLeft      =   795
   ClientTop       =   585
   ClientWidth     =   9735
   Icon            =   "frmMasterAlumni.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9735
   Begin VB.Frame fraFrame 
      Height          =   7215
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   -45
      Width           =   9735
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Hapus Foto"
         Height          =   330
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   6255
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Index           =   23
         Left            =   2790
         TabIndex        =   66
         Top             =   2340
         Width           =   2085
      End
      Begin VB.TextBox txtField 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   330
         Index           =   22
         Left            =   2790
         TabIndex        =   63
         Top             =   1980
         Width           =   2670
      End
      Begin VB.CommandButton cmdAutoNumber 
         Caption         =   "&Auto Number"
         Height          =   330
         Left            =   3375
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   225
         Width           =   1095
      End
      Begin VB.CheckBox chkMenikah 
         Caption         =   "Sudah Menikah"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   60
         Top             =   1710
         Width           =   1590
      End
      Begin VB.TextBox txtField 
         DataField       =   "NamaPerusahaan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   21
         Left            =   1440
         TabIndex        =   59
         Top             =   6705
         Width           =   3525
      End
      Begin VB.TextBox txtField 
         DataField       =   "Pekerjaan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   20
         Left            =   1440
         TabIndex        =   58
         Top             =   6345
         Width           =   2175
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   1
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4725
         Width           =   330
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
         Left            =   5490
         TabIndex        =   43
         Top             =   2025
         Width           =   4020
         Begin VB.TextBox txtField 
            DataField       =   "Telepon5"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   19
            Left            =   1215
            TabIndex        =   53
            Top             =   1800
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon4"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   18
            Left            =   1215
            TabIndex        =   52
            Top             =   1440
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon1"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   17
            Left            =   1215
            TabIndex        =   46
            Top             =   360
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon2"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   16
            Left            =   1215
            TabIndex        =   45
            Top             =   720
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Telepon3"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   15
            Left            =   1215
            TabIndex        =   44
            Top             =   1080
            Width           =   2580
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 5"
            Height          =   195
            Index           =   23
            Left            =   135
            TabIndex        =   51
            Top             =   1845
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 4"
            Height          =   195
            Index           =   22
            Left            =   135
            TabIndex        =   50
            Top             =   1485
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 2"
            Height          =   195
            Index           =   21
            Left            =   135
            TabIndex        =   49
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 3"
            Height          =   195
            Index           =   20
            Left            =   135
            TabIndex        =   48
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lblField 
            Caption         =   "Label Telp 1"
            Height          =   240
            Index           =   19
            Left            =   135
            TabIndex        =   47
            Top             =   405
            Width           =   1005
         End
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
         Left            =   5535
         TabIndex        =   36
         Top             =   315
         Width           =   4020
         Begin VB.TextBox txtField 
            DataField       =   "Email3"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   14
            Left            =   1260
            TabIndex        =   39
            Top             =   1080
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Email2"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   13
            Left            =   1260
            TabIndex        =   38
            Top             =   720
            Width           =   2580
         End
         Begin VB.TextBox txtField 
            DataField       =   "Email1"
            DataSource      =   "dtcMaindata"
            Height          =   330
            Index           =   12
            Left            =   1260
            TabIndex        =   37
            Top             =   360
            Width           =   2580
         End
         Begin VB.Label lblField 
            Caption         =   "Label Email 1"
            Height          =   240
            Index           =   18
            Left            =   135
            TabIndex        =   42
            Top             =   405
            Width           =   1050
         End
         Begin VB.Label lblField 
            Caption         =   "Label Email 3"
            Height          =   195
            Index           =   17
            Left            =   135
            TabIndex        =   41
            Top             =   1125
            Width           =   1050
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
      End
      Begin VB.TextBox txtField 
         DataField       =   "Negara"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   11
         Left            =   1440
         TabIndex        =   35
         Top             =   5985
         Width           =   2175
      End
      Begin VB.TextBox txtField 
         DataField       =   "KodePos"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   10
         Left            =   1440
         TabIndex        =   34
         Top             =   5625
         Width           =   735
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kota"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   9
         Left            =   1440
         TabIndex        =   33
         Top             =   5265
         Width           =   1635
      End
      Begin VB.TextBox txtField 
         DataField       =   "RW"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   8
         Left            =   2475
         TabIndex        =   32
         Top             =   4905
         Width           =   465
      End
      Begin VB.TextBox txtField 
         DataField       =   "RT"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   7
         Left            =   1440
         TabIndex        =   31
         Top             =   4905
         Width           =   465
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kecamatan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   6
         Left            =   1440
         TabIndex        =   30
         Top             =   4545
         Width           =   1635
      End
      Begin VB.TextBox txtField 
         DataField       =   "Kelurahan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   5
         Left            =   1440
         TabIndex        =   29
         Top             =   4185
         Width           =   1635
      End
      Begin VB.TextBox txtField 
         DataField       =   "Alamat"
         DataSource      =   "dtcMaindata"
         Height          =   465
         Index           =   4
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   3690
         Width           =   3570
      End
      Begin MSMask.MaskEdBox mskField 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   27
         Top             =   3330
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtField 
         DataField       =   "TempatLahir"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   3
         Left            =   1440
         TabIndex        =   26
         Top             =   2970
         Width           =   1635
      End
      Begin VB.Frame fraFrame 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   1395
         TabIndex        =   23
         Top             =   2700
         Width           =   1905
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Wanita"
            Height          =   195
            Index           =   1
            Left            =   900
            TabIndex        =   25
            Top             =   0
            Width           =   870
         End
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Pria"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   24
            Top             =   0
            Value           =   -1  'True
            Width           =   690
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboJurusan 
         Bindings        =   "frmMasterAlumni.frx":1E72
         DataField       =   "Jurusan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Left            =   3870
         TabIndex        =   22
         Top             =   1305
         Width           =   960
         DataFieldList   =   "KodeJurusan"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorEven   =   0
         BackColorOdd    =   12648384
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   3200
         Columns(0).Caption=   "KodeJurusan"
         Columns(0).Name =   "KodeJurusan"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "KodeJurusan"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "KodeFakultas"
         Columns(1).Name =   "KodeFakultas"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "KodeFakultas"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   5583
         Columns(2).Caption=   "NamaJurusan"
         Columns(2).Name =   "NamaJurusan"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "NamaJurusan"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   1693
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.ComboBox cboField 
         DataField       =   "Angkatan"
         DataSource      =   "dtcMaindata"
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Text            =   "cboField"
         Top             =   1305
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         DataField       =   "NamaPanggilan"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   20
         Top             =   945
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         DataField       =   "NamaLengkap"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   19
         Top             =   585
         Width           =   3480
      End
      Begin VB.TextBox txtField 
         DataField       =   "NoAlumni"
         DataSource      =   "dtcMaindata"
         Height          =   330
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrompt 
         Height          =   330
         Index           =   0
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   330
      End
      Begin VB.Label lblField 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "LastUpdate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, dd/mmm/yyyy  hh:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "dtcMaindata"
         Height          =   285
         Index           =   31
         Left            =   6840
         TabIndex        =   69
         Top             =   6705
         Width           =   2625
      End
      Begin VB.Label lblField 
         Caption         =   "Perubahan Terakhir"
         Height          =   195
         Index           =   30
         Left            =   5355
         TabIndex        =   68
         Top             =   6750
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Panggilan Pasangan"
         Height          =   195
         Index           =   29
         Left            =   720
         TabIndex        =   65
         Top             =   2340
         Width           =   1950
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Lengkap Pasangan"
         Height          =   195
         Index           =   28
         Left            =   720
         TabIndex        =   64
         Top             =   1980
         Width           =   1950
      End
      Begin VB.Label lblField 
         Caption         =   "Status Menikah"
         Height          =   195
         Index           =   27
         Left            =   135
         TabIndex        =   61
         Top             =   1665
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Perusahaan"
         Height          =   195
         Index           =   26
         Left            =   180
         TabIndex        =   57
         Top             =   6705
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Pekerjaan"
         Height          =   195
         Index           =   25
         Left            =   180
         TabIndex        =   56
         Top             =   6345
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Foto"
         Height          =   195
         Index           =   24
         Left            =   6885
         TabIndex        =   55
         Top             =   4770
         Width           =   465
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   7470
         Stretch         =   -1  'True
         Top             =   4725
         Width           =   1455
      End
      Begin VB.Label lblField 
         Caption         =   "Kota"
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   18
         Top             =   5265
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Negara"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   17
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Kode Pos"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   16
         Top             =   5625
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "RW"
         Height          =   195
         Index           =   12
         Left            =   2115
         TabIndex        =   15
         Top             =   4905
         Width           =   420
      End
      Begin VB.Label lblField 
         Caption         =   "RT"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   14
         Top             =   4905
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Kecamatan"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   13
         Top             =   4590
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Kelurahan"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   12
         Top             =   4230
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Alamat"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   11
         Top             =   3645
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Tanggal Lahir"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   10
         Top             =   3330
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Tempat Lahir"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   9
         Top             =   3015
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Jenis Kelamin"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   8
         Top             =   2700
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Jurusan"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   7
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label lblField 
         Caption         =   "Angkatan"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   1305
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Panggilan"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "Nama Lengkap"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label lblField 
         Caption         =   "No. Alumni"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc dtcMaindata 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6855
      Width           =   9735
      _ExtentX        =   17171
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
   Begin MSAdodcLib.Adodc dtcComboJurusan 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6525
      Width           =   9735
      _ExtentX        =   17171
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
Attribute VB_Name = "frmMasterAlumni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ITransaction

Private bInternalOperation As Boolean
Private bDataChanged As Boolean

Private sFotoFileName As String
Private bIsAutoNumberFetched As Boolean
Private oAutoNum As CAutoNumber

Private Sub cboField_Change()
    '// Beginning of template
    If cboField.DataChanged Then DataChanged = True
    '//
    
    'Your code here...
End Sub

Private Sub cboField_Click()
    If IsBOFEOF(dtcMaindata) Then
        cboField.ListIndex = -1
        Exit Sub
    End If
    
    '// Beginning of template
    If cboField.DataChanged Then DataChanged = True
    '//
End Sub

Private Sub cboJurusan_Change()
    '// Beginning of template
    If cboJurusan.DataChanged Then DataChanged = True
    '//
    
    'Your code here
End Sub

Private Sub cboJurusan_Click()
    '// Beginning of template
    If cboJurusan.DataChanged Then DataChanged = True
    '//
    
    'Your code here...
End Sub

Private Sub cboJurusan_DropDown()
    '// Beginning of template
    If IsBOFEOF(dtcMaindata) Then
        cboJurusan.DroppedDown = False
        Exit Sub
    End If
    '//
    
    'Your code here...
End Sub

Private Sub chkMenikah_Click()
    '// Beginning of template
    If IsBOFEOF(dtcMaindata) Then
        chkMenikah.Value = vbUnchecked
        Exit Sub
    End If
    '//
    
    If bInternalOperation Then Exit Sub
    
    'Your code here...
    If chkMenikah.Value = vbChecked Then
        txtField(22).BackColor = vbWindowBackground
        txtField(22).enabled = True
        txtField(23).BackColor = vbWindowBackground
        txtField(23).enabled = True
    Else
        txtField(22).enabled = False
        txtField(22).BackColor = vbButtonFace
        txtField(22).Text = BLANK
        txtField(23).enabled = False
        txtField(23).BackColor = vbButtonFace
        txtField(23).Text = BLANK
    End If
    DataChanged = True
End Sub

Private Sub cmdAutoNumber_Click()
    If DataChanged Then
        If Trim$(txtField(0).Text) = BLANK Then
            txtField(0).Text = oAutoNum.FetchAutoNumber(MASTER_ALUMNI, BLANK, dtcMaindata.Recordset("NoAlumni").DefinedSize - Len(Parameter![PrefixNoAlumni]))
            Call oAutoNum.IncrementAutoNumber(MASTER_ALUMNI)
            bIsAutoNumberFetched = True
        End If
    End If
End Sub

Private Sub cmdButton_Click()
    '// Beginning of template
    If IsBOFEOF(dtcMaindata) Then Exit Sub
    '//
    
    'Your code here...
    Set imgFoto.Picture = LoadPicture()
    sFotoFileName = BLANK
    dtcMaindata.Recordset![Foto] = BLANK
    DataChanged = True
End Sub

Private Sub cmdPrompt_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Select Case Index
        Case 0
            If DataChanged Then Exit Sub
            sSQL = "SELECT NoAlumni, NamaLengkap FROM Alumni ORDER BY NoAlumni"
            If oDialog.ShowPrompt(ConnectString, sSQL, "Prompt No. Alumni", _
                Array("No. Alumni", "Nama"), txtField(Index).Text) = DialogAnswerOK Then
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
            bInternalOperation = True
            mskField.PromptInclude = False
            mskField.Text = Format$(pRecordset![TanggalLahir], "dd/mm/yyyy")
            mskField.PromptInclude = True
            bInternalOperation = False
            
            bInternalOperation = True
            If UCase$(pRecordset![JenisKelamin] = "P") Then
                
                optJenisKelamin(0).Value = True
            Else
                optJenisKelamin(1).Value = True
            End If
            bInternalOperation = False
            
            bInternalOperation = True
            If CBool(pRecordset![StatusMenikah]) Then chkMenikah.Value = vbChecked Else chkMenikah.Value = vbUnchecked
            bInternalOperation = False
            
            If Trim$(dtcMaindata.Recordset![Foto] & "") <> BLANK Then
                Set imgFoto.Picture = LoadPicture(dtcMaindata.Recordset![Foto])
            Else
                Set imgFoto.Picture = LoadPicture()
            End If
        End If
    Else
        bInternalOperation = True
        mskField.PromptInclude = False
        mskField.Text = BLANK
        mskField.PromptInclude = True
        bInternalOperation = False
        chkMenikah.Value = vbUnchecked
        optJenisKelamin(0).Value = False
        optJenisKelamin(1).Value = False
        Set imgFoto.Picture = LoadPicture()
    End If
    bInternalOperation = False
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

Private Sub Form_Load()
    Dim Control As Control
    
    '// Beginning of Template
    Call LoadPosition(Me, INIPath)
    Call dtcMaindata_Refresh(BLANK)
    '//
    
    
    'Your code here...
    
    Set oAutoNum = New CAutoNumber
    Set oAutoNum.DatabaseConnection = MainDB
    
    With cboField
        .Clear
        For iLoop = 1985 To Year(Now)
            .AddItem CStr(iLoop)
        Next
    End With
    
    For iLoop = cmdPrompt.LBound To cmdPrompt.UBound
        Set cmdPrompt(iLoop).Picture = oPromptIcon
    Next
    
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

Public Sub dtcMaindata_Refresh(ByVal NoAlumni As String)
    With dtcMaindata
        .ConnectionString = ConnectString
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeout
        .LockType = adLockBatchOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .RecordSource = "SELECT * FROM Alumni WHERE NoAlumni = '" & NoAlumni & "'"
        .Refresh
    End With
    
    With dtcComboJurusan
        .ConnectionString = ConnectString
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeout
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .RecordSource = "SELECT * FROM Jurusan ORDER BY KodeJurusan"
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
    DataIsValid = False
    
    If dtcMaindata.Recordset.EditMode = adEditAdd Then
        If Trim$(txtField(0).Text) = BLANK Then
            MsgBox lblField(0).Caption & " tidak boleh dikosongkan", vbInformation, Caption
            Exit Property
        Else
            If IsValidFieldValue(MainDB, "Alumni", "NoAlumni = '" & txtField(0).Text & "'") Then
                MsgBox lblField(0).Caption & " " & txtField(0).Text & " sudah dipakai.", vbInformation, Caption
                Exit Property
            End If
        End If
    Else
        If Trim$(txtField(0).Text) = BLANK Then
            MsgBox lblField(0).Caption & " tidak boleh dikosongkan", vbInformation, Caption
            Exit Property
        Else
            If Trim$(txtField(0).Text) <> Trim$(dtcMaindata.Recordset![NoAlumni].OriginalValue & "") Then
                If IsValidFieldValue(MainDB, "Alumni", "NoAlumni = '" & txtField(0).Text & "'") Then
                    MsgBox lblField(0).Caption & " " & txtField(0).Text & " sudah dipakai.", vbInformation, Caption
                    Exit Property
                End If
            End If
        End If
    End If

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
    optJenisKelamin(0).Value = True

    bIsAutoNumberFetched = False
    dtcMaindata.Recordset.AddNew
    txtField(0).SetFocus
    DataChanged = True
End Sub

Public Sub ITransaction_MasterCancel()
    dtcMaindata.Recordset.Requery
    If bIsAutoNumberFetched Then
        Call oAutoNum.DecrementAutoNumber(MASTER_ALUMNI)
    End If
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
    frmLapDaftarAlumni.Show
    frmLapDaftarAlumni.ZOrder ZOrderConstants.vbBringToFront
End Sub

Public Sub ITransaction_MasterRefresh()
    Call ShowStatusBar("Reloading record from database...")
    Call dtcMaindata_Refresh(txtField(0).Text)
    Call ShowStatusBar("RESET")
End Sub

Public Sub ITransaction_MasterSave()
    Dim lRecordsAffected As Long
    Dim dtLastUpdate As Date
    
    On Error GoTo ErrHandler
    
    If Not DataIsValid Then
        Exit Sub
    End If
    
    If Not IsDate(Format$(mskField.Text, "dd/mm/yyyy")) Then
        dtcMaindata.Recordset![TanggalLahir] = Null
    Else
        dtcMaindata.Recordset![TanggalLahir] = Format$(mskField.Text, "dd/mm/yyyy")
    End If
    
    Select Case True
        Case optJenisKelamin(0).Value
            dtcMaindata.Recordset![JenisKelamin] = "P"
        Case Else
            dtcMaindata.Recordset![JenisKelamin] = "W"
    End Select
    dtcMaindata.Recordset![StatusMenikah] = chkMenikah.Value = vbChecked
    dtcMaindata.Recordset![Foto] = sFotoFileName
        
    dtcMaindata.Recordset.Move 0
    dtcMaindata.Recordset.UpdateBatch
    
    If chkMenikah.Value = vbChecked Then
        MainDB.BeginTrans
        sSQL = "UPDATE Pasangan SET NamaLengkapPasangan = '" & txtField(22).Text & "', " & _
               "NamaPanggilanPasangan = '" & txtField(23).Text & "' " & _
               "WHERE NoAlumni = '" & txtField(0).Text & "'"
        MainDB.Execute sSQL, lRecordsAffected, adCmdText
        If lRecordsAffected = 0 Then
            sSQL = "INSERT INTO Pasangan VALUES ('" & txtField(0).Text & "'," & _
                                              "'" & txtField(22).Text & "'," & _
                                              "'" & txtField(23).Text & "')"
            MainDB.Execute sSQL, lRecordsAffected, adCmdText
        End If
        MainDB.CommitTrans
        Call RefreshDatabaseCache(MainDB)
    Else
        MainDB.BeginTrans
        sSQL = "DELETE FROM Pasangan WHERE NoAlumni = '" & txtField(0).Text & "'"
        MainDB.Execute sSQL, lRecordsAffected, adCmdText
        MainDB.CommitTrans
        Call RefreshDatabaseCache(MainDB)
    End If
        
    MainDB.BeginTrans
    sSQL = "UPDATE Alumni SET LastUpdate=Now WHERE NoAlumni = '" & txtField(0).Text & "'"
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
    
    If IsBOFEOF(dtcMaindata) Then
        Exit Sub
    End If
    
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
    Select Case Index
        Case 15 To 19, 12 To 14
            KeyAscii = NumberOnly(KeyAscii)
    End Select
End Sub
