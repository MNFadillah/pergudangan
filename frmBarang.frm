VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBarang 
   Caption         =   "Data Barang"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   12465
   Begin VB.Frame Data 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdlast 
         Caption         =   ">>"
         Height          =   615
         Left            =   7200
         TabIndex        =   17
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   ">"
         Height          =   615
         Left            =   6240
         TabIndex        =   16
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "<"
         Height          =   615
         Left            =   5280
         TabIndex        =   15
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   "<<"
         Height          =   615
         Left            =   4320
         TabIndex        =   14
         Top             =   3960
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dgBarang 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6165
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txStok 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   2640
         TabIndex        =   11
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   1440
         TabIndex        =   10
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txdesc 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txnama 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Masukkan Nama Barang"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stok"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Deskripsi"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama Barang"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID Barang"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dataBarang As ADODB.Recordset
Dim isEmpty As Boolean

Private Sub cmdAdd_Click()
checkEmpty
If isEmpty = False Then
    On Error GoTo addError
    Dim tgl As Date
    tgl = Format(Now, "dd-MM-yy")
    query = "insert into barang values ('', '" & txnama.Text & "', '" & txdesc.Text & "', '" & tgl & "','" & tgl & "', '" & txStok.Text & "')"
    'query = "insert into barang values ('', '" & txnama.Text & "', '" & txdesc.Text & "',,)"
    'MsgBox tgl
    conn.Execute query
    updateDG
    MsgBox "Berhasil menambahkan data", vbInformation, "Sukses"
    kosongkan
    Exit Sub
addError:
    MsgBox "Gagal menambahkan data " & Err.Description, vbInformation
Else
    MsgBox "Pilih data terlebih dahulu", vbCritical, "Warning"
End If
End Sub

Private Sub cmdChange_Click()
checkEmpty
If isEmpty = False Then
    On Error GoTo addError
    query = "update barang set nama = '" & txnama.Text & "', deskripsi = '" & txdesc.Text & "', stok = " & txStok.Text & " where id = " & txid.Text & ""
    conn.Execute query
    updateDG
    MsgBox "Berhasil mengubah data", vbInformation, "Sukses"
    kosongkan
    Exit Sub
addError:
    MsgBox "Gagal mengubah data " & Err.Description, vbInformation
Else
    MsgBox "Isi data dengan lengkap ", vbCritical, "Warning"
End If
End Sub

Private Sub cmdDelete_Click()
checkEmpty
If isEmpty = False Then
    On Error GoTo addError
    query = "delete from barang where id = " & txid.Text & ""
    conn.Execute query
    updateDG
    MsgBox "Berhasil menghapus data", vbInformation, "Sukses"
    kosongkan
    Exit Sub
addError:
    MsgBox "Gagal menghapus data " & Err.Description, vbInformation
Else
    MsgBox "Pilih data terlebih dahulu", vbCritical, "Warning"
End If
End Sub

Private Sub cmdfirst_Click()
    If Not dataBarang.BOF Then
        dataBarang.MoveFirst
        isitext
    End If
End Sub

Private Sub cmdlast_Click()
If Not dataBarang.AbsolutePosition = dataBarang.RecordCount - 1 Then
        dataBarang.MoveLast
        isitext
    Else
        'MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub cmdnext_Click()
If Not dataBarang.AbsolutePosition = dataBarang.RecordCount Then
        dataBarang.MoveNext
        isitext
    Else
        MsgBox "Sudah data yang paling akhir", vbInformation, "Information"
    End If
End Sub

Private Sub cmdprev_Click()
    If Not dataBarang.AbsolutePosition = 1 Then
        dataBarang.MovePrevious
        isitext
    Else
        MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub


Private Sub dgBarang_Click()
    isitext
End Sub

Private Sub Form_Load()
    updateDG
End Sub

Private Sub updateDG()
    Set dataBarang = New ADODB.Recordset
    dataBarang.Open "select id, nama, deskripsi, stok from barang", conn
    Set dgBarang.DataSource = dataBarang
End Sub

Private Sub checkEmpty()
    If txnama.Text = "" Or txdesc.Text = "" Or txStok.Text = "" Then
        isEmpty = True
    Else
        isEmpty = False
    End If
End Sub

Private Sub kosongkan()
    txid.Text = ""
    txnama.Text = ""
    txdesc.Text = ""
    txStok.Text = ""
End Sub

Private Sub isitext()
    txid = dgBarang.Columns(0).Text
    txnama = dgBarang.Columns(1).Text
    txdesc = dgBarang.Columns(2).Text
    txStok = dgBarang.Columns(3).Text
End Sub
