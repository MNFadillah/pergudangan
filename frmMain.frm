VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu data 
      Caption         =   "Data"
      Begin VB.Menu barang 
         Caption         =   "Barang"
      End
      Begin VB.Menu barang_in 
         Caption         =   "Barang Masuk"
      End
      Begin VB.Menu barang_out 
         Caption         =   "Barang Keluar"
      End
      Begin VB.Menu user 
         Caption         =   "User"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu transaksi_barang_in 
         Caption         =   "Tambah Barang Masuk"
      End
      Begin VB.Menu transaksi_barang_out 
         Caption         =   "Tambah Barang Keluar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub barang_Click()
frmBarang.Show
End Sub

Private Sub barang_in_Click()
FrmBarangIn.Show
End Sub

Private Sub barang_out_Click()
frmBarangOut.Show
End Sub

Private Sub close_Click()
MsgBox "Bye bye", vbInformation, "Exit"
End
End Sub

Private Sub transaksi_barang_in_Click()
frmAddBarangIn.Show
End Sub

Private Sub transaksi_barang_out_Click()
frmAddBarangOut.Show
End Sub
