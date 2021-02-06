VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------
'
'  ITransaction.cls
'
'
'  Class Module ITransaction ini sebenarnya merupakan sebuah interface yang harus diimplementasikan oleh setiap
'  form yang dapat melakukan transaksi penambahan record (add new), penyimpanan (save), penghapusan (delete),
'  pembatalan perubahan (cancel), pencetakan (print), dan pengambilan record kembali (refresh).
'
'  Setiap form yang dapat melakukan transaksi di atas harus memberikan coding sebagai berikut untuk menggunakan
'  interface ini.
'
'                  Implements ITransaction
'
'  yang artinya form tersebut harus mengimplementasikan semua sub/fungsi pada interface ITransaction ini.
'
'
'  (C)Paulus Iman, November 2003-Januari 2005
'  Created exclusively for Persekutuan Alumni Kristen Univ. Bina Nusantara
'
'------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Sub MasterAddNew()
    'Implements to add new record here...
End Sub

Public Sub MasterSave()
    'Implements to save record here...
End Sub

Public Sub MasterDelete()
    'Implements to delete existing record here...
End Sub

Public Sub MasterCancel()
    'Implements to cancel changes of an existing record here...
End Sub

Public Sub MasterPrint()
    'Implements to print record here...
End Sub

Public Sub MasterRefresh()
    'Implements to refresh loaded record here...
End Sub
