VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmShowPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prompt"
   ClientHeight    =   4350
   ClientLeft      =   1650
   ClientTop       =   1965
   ClientWidth     =   6390
   Icon            =   "frmShowPrompt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   4380
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   6360
      Begin MSDataGridLib.DataGrid grdGrid 
         Bindings        =   "frmShowPrompt.frx":0442
         CausesValidation=   0   'False
         Height          =   2535
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
         RowDividerStyle =   4
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
               LCID            =   1057
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
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdShowPrompt 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   3060
         TabIndex        =   11
         Top             =   3825
         Width           =   1725
      End
      Begin VB.CommandButton cmdShowPrompt 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   1
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3825
         Width           =   1725
      End
      Begin TabDlg.SSTab SSTab1 
         CausesValidation=   0   'False
         Height          =   1005
         Left            =   90
         TabIndex        =   2
         Top             =   2745
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   1773
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Find"
         TabPicture(0)   =   "frmShowPrompt.frx":045A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cboField"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cboOp"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdFind"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Sort"
         TabPicture(1)   =   "frmShowPrompt.frx":0476
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdSort"
         Tab(1).Control(1)=   "cboSort"
         Tab(1).Control(2)=   "cboFieldSort"
         Tab(1).ControlCount=   3
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   3195
            TabIndex        =   9
            Top             =   180
            Width           =   1725
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   330
            Left            =   4995
            Picture         =   "frmShowPrompt.frx":0492
            TabIndex        =   8
            Top             =   165
            Width           =   1050
         End
         Begin VB.ComboBox cboFieldSort 
            Height          =   315
            Left            =   -74820
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   180
            Width           =   2715
         End
         Begin VB.ComboBox cboOp 
            Height          =   315
            ItemData        =   "frmShowPrompt.frx":0594
            Left            =   1755
            List            =   "frmShowPrompt.frx":059E
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   180
            Width           =   1320
         End
         Begin VB.ComboBox cboSort 
            Height          =   315
            ItemData        =   "frmShowPrompt.frx":05AF
            Left            =   -72075
            List            =   "frmShowPrompt.frx":05B9
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton cmdSort 
            Caption         =   "&Sort"
            Height          =   330
            Left            =   -70545
            Picture         =   "frmShowPrompt.frx":05D4
            TabIndex        =   4
            Top             =   180
            Width           =   1545
         End
         Begin VB.ComboBox cboField 
            Height          =   315
            ItemData        =   "frmShowPrompt.frx":06D6
            Left            =   135
            List            =   "frmShowPrompt.frx":06E0
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   180
            Width           =   1545
         End
      End
   End
   Begin MSAdodcLib.Adodc dtcPrompt 
      Height          =   2535
      Left            =   5715
      Top             =   -90
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   4471
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
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
      Orientation     =   1
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
Attribute VB_Name = "frmShowPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oRefDialog As CDialog
Private m_ArrayCaption As Variant
Private m_ArrayFields As Variant

Friend Property Set ReferenceDialog(dlg As CDialog)
    Set oRefDialog = dlg
End Property

Friend Property Let ArrayCaption(ArrayCaption As Variant)
    m_ArrayCaption = ArrayCaption
End Property

Friend Property Let ArrayFields(ArrayFields As Variant)
    m_ArrayFields = ArrayFields
End Property

Private Sub cmdFind_Click()
    On Error Resume Next
    dtcPrompt.Recordset.Filter = adFilterNone
    dtcPrompt.Recordset.Requery
    Select Case cboOp.ListIndex
        Case 0 'Match
            dtcPrompt.Recordset.Filter = m_ArrayFields(cboField.ListIndex) & ">='" & Text1.text & "'"
        Case 1 'Like
            dtcPrompt.Recordset.Filter = m_ArrayFields(cboField.ListIndex) & " LIKE '*" & Text1.text & "*'"
    End Select
    
    For iLoop = LBound(m_ArrayCaption) To UBound(m_ArrayCaption)
        grdGrid.Columns(iLoop).Caption = m_ArrayCaption(iLoop)
    Next
    
    cmdShowPrompt(1).Default = True
End Sub

Private Sub cmdShowPrompt_Click(Index As Integer)
    On Error Resume Next
    Dim Rows() As Variant
    Select Case Index
        Case 1   'OK
            ReDim Rows(grdGrid.Columns.Count)
            For iLoop = 0 To grdGrid.Columns.Count - 1
               Rows(iLoop) = grdGrid.Columns(iLoop).text
            Next
            oRefDialog.PromptAnswer Rows, DialogAnswerOK
        Case 0  'Cancel
            oRefDialog.PromptAnswer Empty, DialogAnswerCancel
    End Select
    'Interaction.SaveSetting "X_Utilities", "ShowPrompt", "DefaultFieldIndex", cboField.ListIndex
    Erase m_ArrayFields
    Erase m_ArrayCaption
    Me.Hide
    Unload Me
End Sub

Private Sub cmdSort_Click()
    On Error Resume Next
    Select Case cboSort.ListIndex
        Case 0
            dtcPrompt.Recordset.Sort = m_ArrayFields(cboFieldSort.ListIndex) & " ASC"
        Case 1
            dtcPrompt.Recordset.Sort = m_ArrayFields(cboFieldSort.ListIndex) & " DESC"
    End Select
End Sub

Private Sub grdGrid_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    dtcPrompt.Recordset.Sort = grdGrid.Columns(ColIndex).DataField
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.text)
    cmdFind.Default = True
End Sub

Private Sub Text1_LostFocus()
    cmdShowPrompt(1).Default = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And (Shift = vbAltMask Or Shift = vbCtrlMask) Then
        Call cmdShowPrompt_Click(0)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call cmdShowPrompt_Click(0)
    End If
End Sub
