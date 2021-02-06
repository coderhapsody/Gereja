VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------
'
' CReport.cls
'
' This class contains wrapped functions of Crystal Report Print Engine (CRPE). PAK Binus Program does not use
' Crystal Report Smart Viewer (CRViewer) to preview its reports since I cannot encapsulate CRViewer component into
' a class and hide its unusable attributes. In addition, CRPE32 provides version independent of Crystal Report.
' You can use Crystal Report 7.x or 8.x without any modification of VB codes.
'
' Before utilizing this class, you have to copy CRPE32.DLL to %SystemRoot%\System32.
' See Crystal Report Developer's Help for more information about CRPE.
'
'
' (C)Paulus Iman, November 2003-Januari 2005
' Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------

Option Explicit
Private m_PrintJobHandle As Integer
Private m_TableLocation As PETableLocation
Private m_LogOnInfo As PELogOnInfo

Public Enum EnumDestinationDevice
    ToScreen
    ToPrinter
    ToFile
End Enum

Public Enum EnumOutputFileFormat
    CharSeparated = PE_FT_CHARSEPARATED
    CommaSeparatedValues = PE_FT_CSV
    PlainText = PE_FT_TEXT
    TabFormatted = PE_FT_TABFORMATTED
    TabSeparated = PE_FT_TABSEPARATED
End Enum

Public Enum EnumSortDirection
    Ascending = PE_SF_ASCENDING
    Descending = PE_SF_DESCENDING
End Enum

Public Function SetDatabaseLocation(ByVal TableIndex As Integer, ByVal DatabasePath As String) As Boolean
    With m_TableLocation
        .Location = DatabasePath & vbNullChar
        .StructSize = Len(m_TableLocation)
    End With
    SetDatabaseLocation = CBool(PESetNthTableLocation(m_PrintJobHandle, TableIndex, m_TableLocation))
End Function

Public Function SetRecordselection(ByVal SelectionCriteria As String) As Boolean
    SetRecordselection = CBool(PESetSelectionFormula(m_PrintJobHandle, SelectionCriteria))
End Function

Public Function SetFormula(ByVal formulaName As String, ByVal FormulaValue As String) As Boolean
    SetFormula = CBool(PESetFormula(m_PrintJobHandle, formulaName, FormulaValue))
End Function

Public Function SetSortField(ByVal SortNumber As Integer, ByVal fieldName As String, ByVal SortDirection As EnumSortDirection) As Boolean
    SetSortField = CBool(PESetNthSortField(m_PrintJobHandle, SortNumber, fieldName, SortDirection))
End Function

Public Function ShowReportToPrinter() As Boolean
    Dim hResult As Long
    
    hResult = PEOutputToPrinter(m_PrintJobHandle, 1)
    
    ShowReportToPrinter = CBool(PEStartPrintJob(m_PrintJobHandle, 0))
End Function

Public Function ShowReportToWindow(parentWindow As Object, ByVal Title As String) As Boolean
    Dim hResult As Long, lWidth As Long, lHeight As Long
    
    lWidth = CLng(parentWindow.Width / 20)
    lHeight = CLng(parentWindow.height / 20)
    hResult = PEOutputToWindow(m_PrintJobHandle, Title, 0, 0, lWidth, lHeight, 0, parentWindow.hWnd)
    
    ShowReportToWindow = CBool(PEStartPrintJob(m_PrintJobHandle, 0))
End Function

Public Function OpenReport(ByVal ReportPath As String) As Boolean
    Dim hResult As Long
    
    Call PECloseWindow(m_PrintJobHandle)
    hResult = PEClosePrintJob(m_PrintJobHandle)
    m_PrintJobHandle = PEOpenPrintJob(ReportPath)
    OpenReport = CBool(m_PrintJobHandle)
End Function

Private Sub Class_Initialize()
    If Not CBool(PEOpenEngine()) Then
        MsgBox "Unable to initialize Crystal Report module.", vbCritical, "Internal Error"
    End If
End Sub

Private Sub Class_Terminate()
    Call PECloseWindow(m_PrintJobHandle)
    Call PECloseEngine
End Sub
