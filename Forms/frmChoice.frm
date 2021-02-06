VERSION 5.00
Begin VB.Form frmChoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choice"
   ClientHeight    =   3105
   ClientLeft      =   1560
   ClientTop       =   1995
   ClientWidth     =   4965
   Icon            =   "frmChoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   4965
      Begin VB.CommandButton cmdResponse 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   2520
         Width           =   1275
      End
      Begin VB.CommandButton cmdResponse 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   1275
      End
      Begin VB.ListBox lstChoice 
         Height          =   1620
         ItemData        =   "frmChoice.frx":0442
         Left            =   90
         List            =   "frmChoice.frx":0444
         TabIndex        =   1
         Top             =   810
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmChoice.frx":0446
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblDesc 
         Caption         =   "Select an option below :"
         Height          =   375
         Left            =   675
         TabIndex        =   2
         Top             =   270
         Width           =   4065
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ReferenceDialog As CDialog

Friend Property Set ReferenceDialog(ChoiceDialog As CDialog)
    Set m_ReferenceDialog = ChoiceDialog
End Property

Private Sub cmdResponse_Click(Index As Integer)
    Select Case Index
        Case 0 'OK
            m_ReferenceDialog.ChoiceAnswer lstChoice.List(lstChoice.ListIndex), DialogAnswerOK
        Case 1 'Cancel
            m_ReferenceDialog.ChoiceAnswer Empty, DialogAnswerCancel
    End Select
    Unload Me
End Sub

