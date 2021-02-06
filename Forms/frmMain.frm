VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Pusat Informasi Persekutuan Alumni Kristen Universitas Bina Nusantara"
   ClientHeight    =   7455
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   10785
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picReportToolBar 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6435
      Left            =   9090
      ScaleHeight     =   6435
      ScaleWidth      =   1695
      TabIndex        =   7
      Top             =   750
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CheckBox chkSaveSetting 
         Caption         =   "Save Settings"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2745
         Width           =   1365
      End
      Begin VB.CommandButton cmdReportButton 
         Caption         =   "Sort Order"
         Height          =   780
         Index           =   2
         Left            =   90
         Picture         =   "frmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1845
         Width           =   1500
      End
      Begin VB.CommandButton cmdReportButton 
         Caption         =   "Print Setup"
         Height          =   780
         Index           =   1
         Left            =   90
         Picture         =   "frmMain.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   990
         Width           =   1500
      End
      Begin VB.CommandButton cmdReportButton 
         Caption         =   "Preview"
         Height          =   780
         Index           =   0
         Left            =   90
         Picture         =   "frmMain.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1500
      End
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   8415
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6435
      Left            =   0
      ScaleHeight     =   6435
      ScaleWidth      =   1425
      TabIndex        =   1
      Top             =   750
      Width           =   1425
      Begin VB.Label lblLinkField 
         Caption         =   "Alumni"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   450
         MouseIcon       =   "frmMain.frx":1D68
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1395
         Width           =   555
      End
      Begin VB.Image imgSideBar 
         Height          =   1005
         Index           =   0
         Left            =   180
         MouseIcon       =   "frmMain.frx":2072
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":237C
         Stretch         =   -1  'True
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label lblLinkField 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Retret"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   90
         MouseIcon       =   "frmMain.frx":5E8C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   6930
         Width           =   1095
      End
      Begin VB.Image imgSideBar 
         Height          =   1005
         Index           =   3
         Left            =   135
         MouseIcon       =   "frmMain.frx":6196
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":64A0
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   1005
      End
      Begin VB.Label lblLinkField 
         BackStyle       =   0  'Transparent
         Caption         =   "Pembicara"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   225
         MouseIcon       =   "frmMain.frx":AD72
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   4995
         Width           =   870
      End
      Begin VB.Image imgSideBar 
         Height          =   1005
         Index           =   2
         Left            =   135
         MouseIcon       =   "frmMain.frx":B07C
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":B386
         Stretch         =   -1  'True
         Top             =   4005
         Width           =   1005
      End
      Begin VB.Label lblLinkField 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gereja dan Organisasi Kristen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   645
         Index           =   1
         Left            =   135
         MouseIcon       =   "frmMain.frx":EFFD
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   3015
         Width           =   1095
      End
      Begin VB.Image imgSideBar 
         Height          =   1005
         Index           =   1
         Left            =   180
         MouseIcon       =   "frmMain.frx":F307
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":F611
         Stretch         =   -1  'True
         Top             =   2025
         Width           =   1005
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7185
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12242
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "01/09/2014"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   1323
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "ImageListToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tambah"
            Key             =   "ADDNEW"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpan"
            Key             =   "SAVE"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hapus"
            Key             =   "DELETE"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Batal"
            Key             =   "CANCEL"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cetak"
            Key             =   "PRINT"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "REFRESH"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageListToolBar 
         Left            =   9000
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":133DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":325C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":517B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":718EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":90AD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":AFCBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CEEA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":EE092
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10D27C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12C466
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":14B650
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Pendataan"
      Index           =   0
      Begin VB.Menu mnuPendataan 
         Caption         =   "&Alumni"
         Index           =   0
      End
      Begin VB.Menu mnuPendataan 
         Caption         =   "&Gereja dan Organisasi Kristen"
         Index           =   1
      End
      Begin VB.Menu mnuPendataan 
         Caption         =   "&Pembicara"
         Index           =   2
      End
      Begin VB.Menu mnuPendataan 
         Caption         =   "Tempat &Retret"
         Index           =   3
      End
      Begin VB.Menu mnuPendataan 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPendataan 
         Caption         =   "Persekutuan Alumni Kristen &Jakarta"
         Index           =   5
      End
      Begin VB.Menu mnuPendataan 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuKeluar 
         Caption         =   "&Keluar"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Tabel &Master"
      Index           =   1
      Begin VB.Menu mnuMaster 
         Caption         =   "&Fakultas"
         Index           =   0
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "&Jurusan"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Konfigurasi"
      Index           =   2
      Begin VB.Menu mnuKonfigurasi 
         Caption         =   "Pengaturan &Sistem"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Pencetakan"
      Index           =   3
      Begin VB.Menu mnuPencetakan 
         Caption         =   "Daftar &Alumni"
         Index           =   0
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "Daftar &Gereja dan Organisasi Kristen"
         Index           =   1
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "Daftar &Pembicara"
         Index           =   2
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "Daftar Tempat &Retret"
         Index           =   3
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "Daftar PAK&J"
         Index           =   4
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPencetakan 
         Caption         =   "&Kartu PAK Binus"
         Index           =   6
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Utilitas"
      Index           =   4
      Begin VB.Menu mnuUtilitas 
         Caption         =   "&Hapus Data"
         Index           =   0
      End
      Begin VB.Menu mnuUtilitas 
         Caption         =   "&Pemeliharaan Basis Data"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Informasi"
      Index           =   5
      Begin VB.Menu mnuInformasi 
         Caption         =   "&Cara Pemakaian"
         Index           =   0
      End
      Begin VB.Menu mnuInformasi 
         Caption         =   "&Tentang Program PAK Binus"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSaveSetting_Click()
    'On Error Resume Next
    Me.ActiveForm.SaveSettings = True
End Sub

Private Sub cmdReportButton_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0 'Preview
            Call Me.ActiveForm.IReport_PreviewReport
        Case 1 'Setup
        Case 2 'Sort Order
            Call Me.ActiveForm.IReport_SetSortOrder
    End Select
End Sub

Private Sub imgSideBar_Click(Index As Integer)
    Call mnuPendataan_Click(Index)
End Sub

Private Sub MDIForm_Load()
    Show
    
    Call InitializeMainControls(Toolbar, StatusBar)
    Set oDialog = New CDialog
    
    Call ShowStatusBar("Membuka basis data...")
    
    INIPath = App.Path & "\PAKBinus.INI"
    Set MainDB = New ADODB.Connection
    
    With MainDB
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & App.Path & "\database\PAKBinus.mdb"
        .Open
        
        ConnectString = MainDB.ConnectionString
        
        CommandTimeout = CLng(ToNumeric(ReadINIFile("Konfigurasi", "CommandTimeOut", "0", INIPath)))
    End With
    
    Set Parameter = New ADODB.Recordset
    Parameter.Open "Parameter", MainDB, adOpenStatic, adLockReadOnly, adCmdTable
    
    Me.picSideBar.Visible = Parameter![SideBar]
    
    Call ShowStatusBar("RESET")
    Call DisableToolbarButtons
    
    Set oPromptIcon = ImageListToolBar.ListImages(5).Picture
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("Anda yakin ingin keluar dari program ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
        Cancel = True
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim ChildForm As Form
    
    'Clean up resources we used...
    For Each ChildForm In VB.Global.Forms
        If Not ChildForm Is Me Then
            Unload ChildForm
        End If
    Next
    
    If Not Parameter Is Nothing Then
        If Parameter.State <> adStateClosed Then Parameter.Close
    End If

    If Not MainDB Is Nothing Then
        If MainDB.State <> adStateClosed Then MainDB.Close
    End If
    
    Set MainDB = Nothing
    Set Parameter = Nothing
    Set oDialog = Nothing
End Sub

Private Sub mnuInformasi_Click(Index As Integer)
    Select Case Index
        Case 0
        Case 1
            
            frmAbout.Show vbModal, Me
    End Select
End Sub

Private Sub mnuKeluar_Click()
    Unload Me
End Sub


Private Sub mnuKonfigurasi_Click(Index As Integer)
    Dim ChildForm As Form
    
    
    Select Case Index
        Case 0
            
        Case 1
            If VB.Global.Forms.Count > 1 Then
                If MsgBox("Program akan menutup semua jendela yang ada untuk mencegah inkonsistensi data." & vbNewLine & _
                          "Apakah Anda ingin melanjutkan ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
                    Exit Sub
                End If
                
                For Each ChildForm In VB.Global.Forms
                    If Not ChildForm Is frmMain Then
                        Unload ChildForm
                    End If
                Next
                
            End If
                
                
            frmKonfigurasi.Show
            frmKonfigurasi.ZOrder ZOrderConstants.vbBringToFront
    End Select

End Sub

Private Sub mnuMaster_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMasterFakultas.Show
            frmMasterFakultas.ZOrder ZOrderConstants.vbBringToFront
        Case 1
            frmMasterJurusan.Show
            frmMasterJurusan.ZOrder ZOrderConstants.vbBringToFront
    End Select
End Sub

Private Sub mnuPencetakan_Click(Index As Integer)
    Select Case Index
        Case 0
            frmLapDaftarAlumni.Show
            frmLapDaftarAlumni.ZOrder ZOrderConstants.vbBringToFront
        Case 1
            frmLapDaftarGerejaDanOrganisasi.Show
            frmLapDaftarGerejaDanOrganisasi.ZOrder ZOrderConstants.vbBringToFront
        Case 2
            frmLapDaftarPembicara.Show
            frmLapDaftarPembicara.ZOrder ZOrderConstants.vbBringToFront
        Case 3
            frmLapDaftarTempatRetret.Show
            frmLapDaftarTempatRetret.ZOrder ZOrderConstants.vbBringToFront
        Case 4
            frmLapDaftarPAKJ.Show
            frmLapDaftarPAKJ.ZOrder ZOrderConstants.vbBringToFront
    End Select
    
End Sub

Private Sub mnuPendataan_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMasterAlumni.Show
            frmMasterAlumni.ZOrder ZOrderConstants.vbBringToFront
        Case 1
            frmMasterOrganisasi.Show
            frmMasterOrganisasi.ZOrder ZOrderConstants.vbBringToFront
        Case 2
            frmMasterPembicara.Show
            frmMasterPembicara.ZOrder ZOrderConstants.vbBringToFront
        Case 3
            frmMasterTempatRetret.Show
            frmMasterTempatRetret.ZOrder ZOrderConstants.vbBringToFront
        Case 5
            frmMasterPAKJ.Show
            frmMasterPAKJ.ZOrder ZOrderConstants.vbBringToFront
    End Select
End Sub

Private Sub mnuUtilitas_Click(Index As Integer)
    Dim ChildForm As Form
    
    If VB.Global.Forms.Count > 1 Then
        If MsgBox("Program akan menutup semua jendela yang ada untuk mencegah kerusakan data saat proses berlangsung" & vbNewLine & _
                  "Lanjutkan proses ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
            Exit Sub
        End If
        
        For Each ChildForm In VB.Global.Forms
            If Not ChildForm Is frmMain Then
                Unload ChildForm
            End If
        Next
        
    End If
        
    Select Case Index
        Case 0 'Hapus Data
            frmHapusData.Show
            frmHapusData.ZOrder ZOrderConstants.vbBringToFront
        Case 1 'Compact and Repair
            Call ShowStatusBar("Mohon tunggu, proses pemeliharan data sedang dilakukan....")
            Parameter.Close
            MainDB.Close
            Call CompactDatabase(App.Path & "\Database\PAKBinus.mdb")
            Call ShowStatusBar("Membuka basis data...")
            MainDB.Open
            Parameter.Open "Parameter", MainDB, adOpenForwardOnly, adLockReadOnly, adCmdTable
            Call ShowStatusBar("RESET")
    End Select
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Me.ActiveForm Is Nothing Then Exit Sub

    With Me.ActiveForm
        Select Case UCase$(Button.Key)
            Case "ADDNEW"
                Call .ITransaction_MasterAddNew
            Case "SAVE"
                Call .ITransaction_MasterSave
            Case "DELETE"
                Call .ITransaction_MasterDelete
            Case "CANCEL"
                Call .ITransaction_MasterCancel
            Case "PRINT"
                Call .ITransaction_MasterPrint
            Case "REFRESH"
                Call .ITransaction_MasterRefresh
        End Select
    End With
End Sub
