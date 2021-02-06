VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------
'
' CAutoNumber.cls
'
' This class contains functionality of auto numbering. See documentation at each subs/functions in this class
' for more information.
'
'
' (C)Paulus Iman, November 2003-Januari 2005
' Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------
Option Explicit

Private m_DBCon As ADODB.Connection

Private m_Loop As Integer
Private m_SQL As String

Private Const BLANK As String = ""

'to handle unexpected behaviour while using this class
Public Event AutoNumberError(ByVal ErrorCode As EnumErrorAutoNumber, ByVal ErrorDescription As String)

'Error signal provides by this class
Public Enum EnumErrorAutoNumber
    ERR_CONNECTION_NOT_CREATED = vbObjectError + 1
    ERR_DATABASE_FILENAME_NOT_SET = vbObjectError + 2
    ERR_AUTONUMBER_TABLE_NOT_SET = vbObjectError + 3
    ERR_FAIL_UPDATING_AUTONUMBER_TABLE = vbObjectError + 4
    ERR_AUTONUMBER_LENGTH_INVALID = vbObjectError + 5
    ERR_INVALID_ARGUMENT = vbObjectError + 6
    ERR_UNKNOWN = vbObjectError + 99
End Enum

Public Enum EnumAutoNumberType
    MASTER_ALUMNI
    MASTER_PEMBICARA
    MASTER_GEREJA_DAN_ORGANISASI
    MASTER_PAKJ
    MASTER_TEMPAT_RETRET
End Enum

Private Function FetchAutoNumberValue(ByVal AutoNumberType As EnumAutoNumberType) As Long
    Call RefreshDatabaseCache(m_DBCon)
    
    Select Case AutoNumberType
        Case EnumAutoNumberType.MASTER_ALUMNI
            FetchAutoNumberValue = GetScalarValue(m_DBCon, "SELECT CurrentAutoNumber FROM Parameter")
        Case EnumAutoNumberType.MASTER_PEMBICARA
            FetchAutoNumberValue = GetScalarValue(m_DBCon, "SELECT AutoNumberPembicara FROM Parameter")
        Case EnumAutoNumberType.MASTER_GEREJA_DAN_ORGANISASI
            FetchAutoNumberValue = GetScalarValue(m_DBCon, "SELECT AutoNumberOrganisasi FROM Parameter")
        Case EnumAutoNumberType.MASTER_PAKJ
            FetchAutoNumberValue = GetScalarValue(m_DBCon, "SELECT AutoNumberPAKJ FROM Parameter")
        Case EnumAutoNumberType.MASTER_TEMPAT_RETRET
            FetchAutoNumberValue = GetScalarValue(m_DBCon, "SELECT AutoNumberTempatRetret FROM Parameter")
    End Select
End Function

Public Function FetchAutoNumber(ByVal AutoNumberType As EnumAutoNumberType, ByVal Delimiter As String, Optional ByVal Length As Byte = 4) As String
    Dim lCurrentValue As Long
    Dim sFormatNumber As String
    
    If Length > 10 Then
        RaiseEvent AutoNumberError(ERR_AUTONUMBER_LENGTH_INVALID, "Auto Number cannot more than 10 characters")
        Exit Function
    ElseIf Length < 4 Then
        RaiseEvent AutoNumberError(ERR_AUTONUMBER_LENGTH_INVALID, "Auto Number cannot less than 4 characters")
        Exit Function
    End If
    
    lCurrentValue = FetchAutoNumberValue(AutoNumberType)
        
    For m_Loop = 1 To Length
        sFormatNumber = sFormatNumber & "0"
    Next
    
    Select Case AutoNumberType
        Case EnumAutoNumberType.MASTER_ALUMNI
            FetchAutoNumber = GetFieldValue(m_DBCon, "Parameter", "1", "1", "PrefixNoAlumni", True) & Delimiter & Format$(lCurrentValue, sFormatNumber)
        Case EnumAutoNumberType.MASTER_GEREJA_DAN_ORGANISASI
            FetchAutoNumber = GetFieldValue(m_DBCon, "Parameter", "1", "1", "PrefixArsipOrganisasi", True) & Delimiter & Format$(lCurrentValue, sFormatNumber)
        Case EnumAutoNumberType.MASTER_PAKJ
            FetchAutoNumber = GetFieldValue(m_DBCon, "Parameter", "1", "1", "PrefixArsipPAKJ", True) & Delimiter & Format$(lCurrentValue, sFormatNumber)
        Case EnumAutoNumberType.MASTER_PEMBICARA
            FetchAutoNumber = GetFieldValue(m_DBCon, "Parameter", "1", "1", "PrefixArsipPembicara", True) & Delimiter & Format$(lCurrentValue, sFormatNumber)
        Case EnumAutoNumberType.MASTER_TEMPAT_RETRET
            FetchAutoNumber = GetFieldValue(m_DBCon, "Parameter", "1", "1", "PrefixArsipTempatRetret", True) & Delimiter & Format$(lCurrentValue, sFormatNumber)
    End Select
End Function

Public Function DecrementAutoNumber(ByVal AutoNumberType As EnumAutoNumberType) As Boolean
    Dim lRecordsAffected As Long
    
    m_DBCon.BeginTrans
    
    Select Case AutoNumberType
        Case EnumAutoNumberType.MASTER_ALUMNI
            m_SQL = "UPDATE Parameter SET CurrentAutoNumber=CurrentAutoNumber-1 WHERE CurrentAutoNumber > 1 "
        Case EnumAutoNumberType.MASTER_GEREJA_DAN_ORGANISASI
            m_SQL = "UPDATE Parameter SET AutoNumberOrganisasi=AutoNumberOrganisasi-1 WHERE AutoNumberOrganisasi > 1"
        Case EnumAutoNumberType.MASTER_PAKJ
            m_SQL = "UPDATE Parameter SET AutoNumberPAKJ=AutoNumberPAKJ-1 WHERE AutoNumberPAKJ > 1 "
        Case EnumAutoNumberType.MASTER_PEMBICARA
            m_SQL = "UPDATE Parameter SET AutoNumberPembicara=AutoNumberPembicara-1 WHERE AutoNumberPembicara > 1 "
        Case EnumAutoNumberType.MASTER_TEMPAT_RETRET
            m_SQL = "UPDATE Parameter SET AutoNumberTempatRetret=AutoNumberTempatRetret-1 WHERE AutoNumberTempatRetret > 1 "
    End Select
    m_DBCon.Execute m_SQL, lRecordsAffected, adCmdText
    
    m_DBCon.CommitTrans
    Call RefreshDatabaseCache(m_DBCon)
    
    If lRecordsAffected <= 0 Then
        RaiseEvent AutoNumberError(ERR_FAIL_UPDATING_AUTONUMBER_TABLE, "Fail when trying to update autonumber table")
    End If
    
    DecrementAutoNumber = True
End Function

Public Function IncrementAutoNumber(ByVal AutoNumberType As EnumAutoNumberType) As Boolean
    Dim lRecordsAffected As Long

    m_DBCon.BeginTrans

    Select Case AutoNumberType
        Case EnumAutoNumberType.MASTER_ALUMNI
            m_SQL = "UPDATE Parameter SET CurrentAutoNumber=CurrentAutoNumber+1"
        Case EnumAutoNumberType.MASTER_GEREJA_DAN_ORGANISASI
            m_SQL = "UPDATE Parameter SET AutoNumberOrganisasi=AutoNumberOrganisasi+1"
        Case EnumAutoNumberType.MASTER_PAKJ
            m_SQL = "UPDATE Parameter SET AutoNumberPAKJ=AutoNumberPAKJ+1"
        Case EnumAutoNumberType.MASTER_PEMBICARA
            m_SQL = "UPDATE Parameter SET AutoNumberPembicara=AutoNumberPembicara+1"
        Case EnumAutoNumberType.MASTER_TEMPAT_RETRET
            m_SQL = "UPDATE Parameter SET AutoNumberTempatRetret=AutoNumberTempatRetret+1"
    End Select
    m_DBCon.Execute m_SQL, lRecordsAffected, adCmdText
    
    m_DBCon.CommitTrans
    Call RefreshDatabaseCache(m_DBCon)
    
    If lRecordsAffected <= 0 Then
        RaiseEvent AutoNumberError(ERR_FAIL_UPDATING_AUTONUMBER_TABLE, "Fail when trying to update autonumber table")
    End If
    
    IncrementAutoNumber = True
End Function

Public Property Set DatabaseConnection(DBConnection As ADODB.Connection)
    Set m_DBCon = DBConnection
End Property

Public Property Get DatabaseConnection() As ADODB.Connection
    Set DatabaseConnection = m_DBCon
End Property

