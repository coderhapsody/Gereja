VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------
'
'  IReport.cls
'
'
'  Class Module IReport ini sebenarnya merupakan sebuah interface yang harus diimplementasikan oleh setiap
'  form yang berfungsi untuk menampilkan laporan (report).
'
'  Setiap form yang dapat menampilkan laporan (report) harus memberikan coding sebagai berikut untuk menggunakan
'  interface ini.
'
'                  Implements IReport
'
'  yang artinya form tersebut harus mengimplementasikan semua sub/fungsi pada interface IReport ini.
'
'
'  (C)Paulus Iman, November 2003-Januari 2005
'  Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------

Option Explicit

Public Sub PreviewReport()
    'Implements to preview report on screen here...
End Sub

Public Sub SetSortOrder()
    'Implements to change sort order of report here...
End Sub

