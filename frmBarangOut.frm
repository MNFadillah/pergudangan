VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBarangOut 
   Caption         =   "Data Barang Keluar"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin MSDataGridLib.DataGrid databarang 
         Height          =   4455
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7858
         _Version        =   393216
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
      Begin VB.CommandButton cmdlast 
         Caption         =   ">>"
         Height          =   615
         Left            =   4680
         TabIndex        =   20
         Top             =   4800
         Width           =   735
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   ">"
         Height          =   615
         Left            =   3840
         TabIndex        =   19
         Top             =   4800
         Width           =   735
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "<"
         Height          =   615
         Left            =   3000
         TabIndex        =   18
         Top             =   4800
         Width           =   735
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   "<<"
         Height          =   615
         Left            =   2160
         TabIndex        =   17
         Top             =   4800
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmddelete 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   2280
         TabIndex        =   16
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdbrowse2 
         Caption         =   ". . ."
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdbrowse1 
         Caption         =   ". . ."
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox departemen 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4440
         Width           =   3135
      End
      Begin VB.TextBox jumlah 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox userId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox barangId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox id 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DEPARTEMEN"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "JUMLAH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID USER"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID BARANG"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmBarangOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim databarangout As New ADODB.Recordset
Dim isempty As Boolean

Private Sub cmdadd_Click()
checkEmpty
If isempty = False Then
    On Error GoTo addError
    query = "insert into barang_keluar values ('', '" & barangId.Text & "', '" & userId.Text & "', " & jumlah.Text & ", '" & departemen.Text & "', '')"
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

Private Sub cmdbrowse1_Click()
    frmData.table = 2
    frmData.kolom = "barang"
    frmData.Show
End Sub

Private Sub cmdbrowse2_Click()
    frmData.table = 2
    frmData.kolom = "user"
    frmData.Show
End Sub

Private Sub cmddelete_Click()
checkEmpty
If isempty = False Then
    On Error GoTo addError
    query = "delete from barang_keluar where id = " & id.Text & ""
    conn.Execute query
    MsgBox "Berhasil menghapus data", vbInformation, "Sukses"
    updateDG
    kosongkan
    Exit Sub
addError:
    MsgBox "Gagal menghapus data " & Err.Description, vbInformation
Else
    MsgBox "Pilih data terlebih dahulu", vbCritical, "Warning"
End If
End Sub


Private Sub cmdfirst_Click()
If Not databarangout.BOF Then
        databarangout.MoveFirst
        isitext
    End If
End Sub

Private Sub cmdlast_Click()
If Not databarangout.AbsolutePosition = databarangout.RecordCount Then
        databarangout.MoveLast
        isitext
    Else
        'MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub cmdnext_Click()
If Not databarangout.AbsolutePosition = databarangout.RecordCount Then
        databarangout.MoveNext
        isitext
    Else
        MsgBox "Sudah data yang paling akhir", vbInformation, "Information"
    End If
End Sub

Private Sub cmdprev_Click()
If Not databarangout.AbsolutePosition = 1 Then
        databarangout.MovePrevious
        isitext
    Else
        MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub cmdupdate_Click()
checkEmpty
If isempty = False Then
    On Error GoTo addError
    query = "update barang_keluar set id_barang = '" & barangId.Text & "', id_user = '" & userId.Text & "', jumlah = " & jumlah.Text & ", departemen = '" & departemen.Text & "' where id = " & id.Text & ""
    conn.Execute query
    MsgBox "Berhasil mengubah data", vbInformation, "Sukses"
    updateDG
    kosongkan
    Exit Sub
addError:
    MsgBox "Gagal mengubah data " & Err.Description, vbInformation
Else
    MsgBox "Isi data dengan lengkap ", vbCritical, "Warning"
End If
End Sub
Private Sub databarang_Click()
    isitext
End Sub

Private Sub Form_Load()
    updateDG
End Sub

Private Sub isitext()
id.Text = databarang.Columns(0).Text
barangId.Text = databarang.Columns(1).Text
userId.Text = databarang.Columns(2).Text
jumlah.Text = databarang.Columns(3).Text
departemen.Text = databarang.Columns(4).Text
End Sub

Private Sub checkEmpty()
If barangId.Text = "" Or userId.Text = "" Or jumlah.Text = "" Or departemen.Text = "" Then
    isempty = True
Else
    isempty = False
End If
End Sub

Private Sub updateDG()
    Set databarangout = New ADODB.Recordset
    databarangout.Open "select id, id_barang, id_user, jumlah, departemen from barang_keluar", conn
    Set databarang.DataSource = databarangout
End Sub

Private Sub kosongkan()
id.Text = ""
barangId.Text = ""
userId.Text = ""
jumlah.Text = ""
departemen.Text = ""
End Sub
