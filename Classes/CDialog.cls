VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------
'
' CDialog.cls
'
' This class contains wrapped functions of common dialog used in PAK Binus program, so that PAKBinus
' program has consistent view of dialog boxes.
'
'
' (C)Paulus Iman, November 2003-Januari 2005
' Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------
Option Explicit

' Jawaban yang dapat diberikan oleh user untuk setiap dialog yang ditampilkan melalui class ini.
Public Enum DialogAnswerEnum
    DialogAnswerOK              'User memilih button OK.
    DialogAnswerCancel          'User memilih button CANCEL.
End Enum

Private m_PromptAnswers As Variant
Private m_ChoiceAnswer As Variant
Private m_SelectedDate As Variant

Private m_DefaultDatabase As Variant
Private m_DefaultServer As Variant

Private m_ButtonAnswer As DialogAnswerEnum

' Initialize absolute paging size
Private Const PAGE_SIZE As Long = 50

' Sub PromptAnswer() digunakan secara internal untuk mengambil jawaban dari fungsi ShowPrompt()
Friend Sub PromptAnswer(Answer As Variant, ButtonAnswer As DialogAnswerEnum)
    m_PromptAnswers = Answer
    m_ButtonAnswer = ButtonAnswer
End Sub

' Sub ChoiceAnswer() digunakan secara internal untuk mengambil jawaban dari fungsi ShowChoice()
Friend Sub ChoiceAnswer(Answer As Variant, ButtonAnswer As DialogAnswerEnum)
    m_ChoiceAnswer = Answer
    m_ButtonAnswer = ButtonAnswer
End Sub

' Sub DateAnswer() digunakan secara internal untuk mengambil jawaban dari fugnsi ShowDate()
Friend Sub DateAnswer(Answer As Variant, ButtonAnswer As DialogAnswerEnum)
    m_SelectedDate = Answer
    m_ButtonAnswer = ButtonAnswer
End Sub

'
' Fungsi ShowPrompt() dapat digunakan untuk menampilkan prompt.
'
' Paramter Input :
'   ConnectionString    String      Connection String yang dapat diterima OLEDB
'   SQL                 String      Statement SQL untuk mengambil record
'   PromptTitle         String      Judul atas dari prompt
'   ColumnsTitle        Variant     Array untuk mengisi kolom-kolom dari record-record yang diambil oleh statement SQL
'   [InitialText]       String      Teks awal untuk pencarian
'   [GridTitle]         String      Judul atas dari grid
'
' Parameter Output :
'   N/A
'
' Return Value : DialogAnswerEnum
'   DialogAnswerOK      Jika user memilih button OK
'   DialogAnswerCancel  Jika user memilih button Cancel
'
Public Function ShowPrompt(ByVal ConnectionString As String, _
                           ByVal SQL As String, _
                           ByVal PromptTitle As String, _
                           ColumnsTitle As Variant, _
                           Optional ByVal InitialText As String = BLANK, _
                           Optional ByVal GridTitle As String = BLANK) As DialogAnswerEnum
On Error GoTo ErrHandler
Dim ArrayFields As Variant
Dim iDefaultFieldIndex As Integer

    ReDim ArrayFields(UBound(ColumnsTitle) + 1)
    
    Load frmShowPrompt
    With frmShowPrompt
        Set .ReferenceDialog = Me
        .Caption = PromptTitle
        .dtcPrompt.ConnectionString = ConnectionString
        .dtcPrompt.RecordSource = SQL
        .dtcPrompt.Refresh
        .dtcPrompt.Recordset.PageSize = PAGE_SIZE
        .grdGrid.Caption = GridTitle
        .cboField.Clear
        .cboFieldSort.Clear
        For iLoop = 1 To .dtcPrompt.Recordset.Fields.Count
            .grdGrid.Columns(iLoop - 1).Caption = ColumnsTitle(iLoop - 1)
            .cboField.AddItem ColumnsTitle(iLoop - 1)
            .cboFieldSort.AddItem ColumnsTitle(iLoop - 1)
            ArrayFields(iLoop - 1) = .dtcPrompt.Recordset.Fields(iLoop - 1).Name
        Next
        .ArrayCaption = ColumnsTitle
        .ArrayFields = ArrayFields
        .cboOp.ListIndex = 0
        .cboSort.ListIndex = 0
        'iDefaultFieldIndex = CInt(Interaction.GetSetting("X_Utilities", "ShowPrompt", "DefaultFieldIndex", "0"))
        'If iDefaultFieldIndex > .cboField.ListCount Then
        '    .cboField.ListIndex = 0
        'Else
        '    .cboField.ListIndex = iDefaultFieldIndex
        'End If
        .cboField.ListIndex = 0
        .cboFieldSort.ListIndex = .cboField.ListIndex
        .dtcPrompt.Recordset.Sort = ArrayFields(0)
        .Text1.Text = InitialText
        'If InitialText <> BLANK Then
        '    .dtcPrompt.Recordset.Filter = ArrayFields(.cboField.ListIndex) & ">='" & InitialText & "'"
        '    If .dtcPrompt.Recordset.RecordCount = 0 Then
        '        .dtcPrompt.Recordset.Filter = adFilterNone
        '        .dtcPrompt.Recordset.Requery
        '    End If
        'End If
        .Show vbModal
    End With
    ShowPrompt = m_ButtonAnswer
    Exit Function

ErrHandler:
    MsgBox Err.Description, vbCritical, "DLL Internal Error"
'    Resume
End Function
   
'
' Property ColumnValue() dapat digunakan untuk mengambil isi kolom dari record yang ditunjuk pada window ShowPrompt.
'
' Parameter Input :
'   Index       Integer         Indeks kolom yang akan diambil isinya.
'
' Parameter Output :
'   N/A
'
' Return Value : Variant
'   Isi kolom sesuai indeks yang diberikan pada parameter input.
'
Public Property Get ColumnValue(ByVal Index As Integer) As Variant
On Error GoTo ErrHandler
    If Not IsEmpty(m_PromptAnswers) Then
        ColumnValue = m_PromptAnswers(Index)
    End If
    Exit Property
    
ErrHandler:
    MsgBox Err.Description, vbCritical, "DLL Internal Error"
End Property


Public Function ShowChoice(Choices() As String, Optional ByVal Title As String = BLANK) As DialogAnswerEnum
    Dim Choice As Variant
    
    Load frmChoice
    
    If Title <> BLANK Then
        frmChoice.lblDesc.Caption = Title
    End If
    
    Set frmChoice.ReferenceDialog = Me
    For Each Choice In Choices
        frmChoice.lstChoice.AddItem Choice
    Next
    frmChoice.Show vbModal
    ShowChoice = m_ButtonAnswer
End Function

Public Property Get SelectedChoice() As Variant
    SelectedChoice = m_ChoiceAnswer
End Property

Public Function ShowDatePrompt(ByVal InitialDate As Date) As DialogAnswerEnum
    Load frmDatePrompt
    With frmDatePrompt
        Set .ReferenceDialog = Me
        .Calendar.Value = InitialDate
        .Show vbModal
        ShowDatePrompt = m_ButtonAnswer
    End With
End Function

Public Property Get SelectedDate() As Variant
    SelectedDate = Format(m_SelectedDate, "mm/dd/yyyy")
End Property
