Attribute VB_Name = "mdlGlobalVars"
Option Explicit

'---------------------------------------------
'              GLOBAL VARIABLES
'---------------------------------------------

'Used to construct SQL statement
Global sSQL As String

'ADODB Connection to database
Global MainDB As ADODB.Connection

'Main connection string for ADODB.Connection
Global ConnectString As String

'Time out used by ADO to wait command execution (in seconds)
Global CommandTimeout As Long

'Used commonly by For..Next looping
Global iLoop As Integer
Global lLoop As Long

'Path for INI file
Global INIPath As String

'Object for PAKBinus common dialog class
Global oDialog As CDialog

'Recordset for PAKBinus main configuration
Global Parameter As ADODB.Recordset

'Icon for prompt button
Global oPromptIcon As StdPicture


'---------------------------------------------
'                CONSTANTS
'---------------------------------------------
Global Const BLANK As String = ""
