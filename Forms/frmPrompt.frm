VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrompt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prompt"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   Begin MSAdodcLib.Adodc dtcPrompt 
      Height          =   330
      Left            =   4770
      Top             =   2700
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
   Begin TabDlg.SSTab TabUtility 
      Height          =   825
      Left            =   90
      TabIndex        =   3
      Top             =   2295
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1455
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Find"
      TabPicture(0)   =   "frmPrompt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblField(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtField"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFind"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Sort"
      TabPicture(1)   =   "frmPrompt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboField(1)"
      Tab(1).Control(1)=   "cboField(0)"
      Tab(1).Control(2)=   "lblField(1)"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Default         =   -1  'True
         Height          =   330
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Find"
         Top             =   405
         Width           =   735
      End
      Begin VB.ComboBox cboField 
         Height          =   315
         Index           =   1
         ItemData        =   "frmPrompt.frx":0038
         Left            =   -71850
         List            =   "frmPrompt.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   405
         Width           =   1275
      End
      Begin VB.ComboBox cboField 
         Height          =   315
         Index           =   0
         Left            =   -74190
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   405
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Height          =   330
         Left            =   1485
         TabIndex        =   5
         Top             =   405
         Width           =   2130
      End
      Begin VB.Label lblField 
         Caption         =   "Column"
         Height          =   195
         Index           =   1
         Left            =   -74820
         TabIndex        =   7
         Top             =   450
         Width           =   555
      End
      Begin VB.Label lblField 
         Caption         =   "Enter expression :"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   450
         Width           =   1275
      End
   End
   Begin MSDataGridLib.DataGrid grdGrid 
      Bindings        =   "frmPrompt.frx":005D
      Height          =   2130
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3757
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      Picture         =   "frmPrompt.frx":0075
      TabIndex        =   1
      Top             =   585
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      Picture         =   "frmPrompt.frx":1F24F
      TabIndex        =   0
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click()
    Dim sFieldName As String
    
    sFieldName = grdGrid.Columns(grdGrid.Col).DataField
    
    If Not dtcPrompt.Recordset Is Nothing Then
        With dtcPrompt.Recordset
            .Filter = sFieldName & "LIKE >= '" & txtField.Text & "'"
        End With
    End If
    
    cmdOK.Default = True
End Sub
