VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmDatePrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date Prompt"
   ClientHeight    =   3285
   ClientLeft      =   1695
   ClientTop       =   1875
   ClientWidth     =   4590
   Icon            =   "frmDatePrompt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   3300
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   4560
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   2475
         TabIndex        =   3
         Top             =   2790
         Width           =   1410
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   2790
         Width           =   1410
      End
      Begin MSACAL.Calendar Calendar 
         Height          =   2670
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   135
         Width           =   4380
         _Version        =   524288
         _ExtentX        =   7726
         _ExtentY        =   4710
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2004
         Month           =   9
         Day             =   5
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   0   'False
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   0   'False
         TitleFontColor  =   8388608
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDatePrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ReferenceDialog As CDialog

Friend Property Set ReferenceDialog(ChoiceDialog As CDialog)
    Set m_ReferenceDialog = ChoiceDialog
End Property

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
            m_ReferenceDialog.DateAnswer Calendar.Value, DialogAnswerOK
        Case 1
            m_ReferenceDialog.DateAnswer Empty, DialogAnswerCancel
    End Select
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And (Shift = vbAltMask Or Shift = vbCtrlMask) Then
        Call cmdAction_Click(1)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call cmdAction_Click(1)
    End If
End Sub

