VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMasterFakultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabel Master Fakultas"
   ClientHeight    =   3225
   ClientLeft      =   3495
   ClientTop       =   3255
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6285
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdGrid 
      Bindings        =   "frmMasterFakultas.frx":0000
      Height          =   3210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6270
      _Version        =   196617
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      AllowColumnMoving=   0
      AllowColumnSwapping=   0
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2064
      Columns(0).Caption=   "Kode"
      Columns(0).Name =   "KodeFakultas"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "KodeFakultas"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5080
      Columns(1).Caption=   "Nama"
      Columns(1).Name =   "NamaFakultas"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "NamaFakultas"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   11060
      _ExtentY        =   5662
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc dtcMaindata 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   2850
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   661
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
Attribute VB_Name = "frmMasterFakultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bGridIsDeleted As Boolean
Private bInternalOperation As Boolean
Private bDataChanged As Boolean

Implements ITransaction

Private Sub Form_Activate()
    DataChanged = bDataChanged
End Sub

Private Sub Form_Deactivate()
    bDataChanged = DataChanged
End Sub

Private Sub Form_GotFocus()
    DataChanged = bDataChanged
End Sub

Private Sub Form_Load()
    Call LoadPosition(Me, INIPath)
    Call dtcMaindata_Refresh
    DataChanged = False
End Sub

Private Sub dtcMaindata_Refresh()
    With dtcMaindata
        .ConnectionString = ConnectString
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeout
        .RecordSource = "SELECT * FROM Fakultas ORDER BY KodeFakultas"
        .Refresh
    End With
End Sub

Private Sub Form_LostFocus()
    bDataChanged = DataChanged
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If grdGrid.RowChanged Then
        MsgBox "Record harus disimpan sebelum menutup form ini.", vbInformation, Caption
        grdGrid.SetFocus
        grdGrid.Col = grdGrid.Columns("KodeFakultas").Position
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePosition(Me, INIPath)
    Call DisableToolbarButtons
End Sub

Private Sub grdGrid_AfterDelete(RtnDispErrMsg As Integer)
    grdGrid.Col = grdGrid.Columns("KodeFakultas").Position
    bGridIsDeleted = False
End Sub

Private Sub grdGrid_AfterUpdate(RtnDispErrMsg As Integer)
    grdGrid.Col = grdGrid.Columns("KodeFakultas").Position
    bGridIsDeleted = False
    DataChanged = False
End Sub

Private Sub grdGrid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    If grdGrid.RowChanged Or DataChanged Then Exit Sub
    If IsBOFEOF(dtcMaindata) Then Exit Sub
    
    DispPromptMsg = False
    grdGrid.SelBookmarks.Add dtcMaindata.Recordset.Bookmark
    If MsgBox("Hapus record ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
        bGridIsDeleted = False
        Cancel = True
        grdGrid.SelBookmarks.RemoveAll
        grdGrid.SetFocus
        grdGrid.Col = grdGrid.Columns("KodeFakultas").Position
        Exit Sub
    End If
    
    bGridIsDeleted = True
End Sub

Private Sub grdGrid_BeforeUpdate(Cancel As Integer)
    If bGridIsDeleted Then Exit Sub
    If bInternalOperation Then Exit Sub
    
    If Not DataIsValid Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Property Get DataIsValid() As Boolean
    With grdGrid
        If Trim$(.Columns("KodeFakultas").Text) = BLANK Then
            MsgBox .Columns("KodeFakultas").Caption & " tidak boleh dikosongkan.", vbInformation, Caption
            .SetFocus
            .Col = .Columns("KodeFakultas").Position
            Exit Property
        Else
            If .IsAddRow Then
                If IsValidFieldValue(MainDB, "Fakultas", "KodeFakultas = '" & .Columns("KodeFakultas").Text & "'") Then
                    MsgBox .Columns("KodeFakultas").Caption & " " & .Columns("KodeFakultas").Text & " sudah ada.", vbInformation, Caption
                    .SetFocus
                    .Col = .Columns("KodeFakultas").Position
                    Exit Property
                End If
            Else
                If .Columns("KodeFakultas").Text <> dtcMaindata.Recordset!KodeFakultas Then
                    If IsValidFieldValue(MainDB, "Fakultas", "KodeFakultas = '" & .Columns("KodeFakultas").Text & "'") Then
                        MsgBox .Columns("KodeFakultas").Caption & " " & .Columns("KodeFakultas").Text & " sudah ada.", vbInformation, Caption
                        .SetFocus
                        .Col = .Columns("KodeFakultas").Position
                        Exit Property
                    End If
                End If
            End If
        End If
        
        If Trim$(.Columns("NamaFakultas").Text) = BLANK Then
            MsgBox .Columns("NamaFakultas").Caption & " tidak boleh dikosongkan.", vbInformation, Caption
            .SetFocus
            .Col = .Columns("NamaFakultas").Position
            Exit Property
        End If
    End With
    
    DataIsValid = True
End Property

Private Sub grdGrid_Change()
    If Not DataChanged Then DataChanged = True
End Sub

Private Sub grdGrid_HeadClick(ByVal ColIndex As Integer)
    If grdGrid.Columns(ColIndex).DataField = BLANK Then Exit Sub
    
    dtcMaindata.Recordset.Sort = grdGrid.Columns(ColIndex).DataField
End Sub

Private Sub grdGrid_InitColumnProps()
    With grdGrid
        .Columns("KodeFakultas").FieldLen = dtcMaindata.Recordset.Fields("KodeFakultas").DefinedSize
        .Columns("NamaFakultas").FieldLen = dtcMaindata.Recordset.Fields("NamaFakultas").DefinedSize
    End With
End Sub

Private Sub grdGrid_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckKeyPress(KeyAscii)

    If KeyAscii = vbKeyEscape Then
        grdGrid.CancelUpdate
        DataChanged = False
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Public Sub ITransaction_MasterAddNew()

    dtcMaindata.EOFAction = adStayEOF
    If Not dtcMaindata.Recordset Is Nothing And dtcMaindata.Recordset.RecordCount > 0 Then
        dtcMaindata.Recordset.MoveLast
        dtcMaindata.Recordset.MoveNext
    End If
    grdGrid.SetFocus
    grdGrid.Col = grdGrid.Columns("KodeFakultas").Position
    
    DataChanged = True
End Sub

Public Sub ITransaction_MasterCancel()
    dtcMaindata.Refresh
    DataChanged = False
    
    grdGrid.SetFocus
    grdGrid.Col = grdGrid.Columns("KodeFakultas").Position
End Sub

Public Sub ITransaction_MasterDelete()
    If IsBOFEOF(dtcMaindata) Then Exit Sub
    If grdGrid.RowChanged Or DataChanged Then Exit Sub
    
    grdGrid.DeleteSelected
    DataChanged = False
End Sub

Public Sub ITransaction_MasterPrint()

End Sub

Public Sub ITransaction_MasterRefresh()
    If DataChanged Then Exit Sub
    Call dtcMaindata_Refresh
End Sub

Public Sub ITransaction_MasterSave()
    grdGrid.Update
End Sub

Private Property Let DataChanged(ByVal Toggle As Boolean)
    Call AdjustToolbarButton(Not Toggle, Toggle, Not Toggle, Toggle)
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = frmMain.Toolbar.Buttons("SAVE").enabled
End Property
