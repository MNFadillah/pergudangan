VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBarangIn 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command10 
         Caption         =   ">>"
         Height          =   615
         Left            =   5400
         TabIndex        =   21
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   ">"
         Height          =   615
         Left            =   4560
         TabIndex        =   20
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "<"
         Height          =   615
         Left            =   3720
         TabIndex        =   19
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "<<"
         Height          =   615
         Left            =   2880
         TabIndex        =   18
         Top             =   4560
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid dgbarangin 
         Height          =   4215
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7435
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
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command6 
         Caption         =   ". . ."
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   ". . ."
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   2520
         TabIndex        =   14
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   1320
         TabIndex        =   13
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox vendor 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox jumlah 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox userid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox barangid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox id 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "VENDOR"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "JUMLAH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2760
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
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID BARANG"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
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
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmBarangIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim databarangin As New ADODB.Recordset
Dim isEmpty As Boolean

Private Sub Command1_Click()
checkEmpty
If isEmpty = False Then
    On Error GoTo addError
    query = "insert into barang_masuk values ('', '" & barangid.Text & "', '" & userid.Text & "', " & jumlah.Text & ", '" & vendor.Text & "', '')"
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

Private Sub Command10_Click()
If Not databarangin.AbsolutePosition = databarangin.RecordCount Then
        databarangin.MoveLast
        isitext
    Else
        'MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub Command2_Click()
checkEmpty
If isEmpty = False Then
    On Error GoTo addError
    query = "update barang_masuk set id_barang = '" & barangid.Text & "', id_user = '" & userid.Text & "', jumlah = " & jumlah.Text & ", vendor = '" & vendor.Text & "' where id = " & id.Text & ""
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

Private Sub Command3_Click()
checkEmpty
If isEmpty = False Then
    On Error GoTo addError
    query = "delete from barang_masuk where id = " & id.Text & ""
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

Private Sub Command5_Click()
    frmData.table = 1
    frmData.kolom = "barang"
    frmData.Show
End Sub

Private Sub Command6_Click()
    frmData.table = 1
    frmData.kolom = "user"
    frmData.Show
End Sub

Private Sub Command7_Click()
If Not databarangin.BOF Then
        databarangin.MoveFirst
        isitext
    End If
End Sub

Private Sub Command8_Click()
If Not databarangin.AbsolutePosition = 1 Then
        databarangin.MovePrevious
        isitext
    Else
        MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub Command9_Click()
If Not databarangin.AbsolutePosition = databarangin.RecordCount Then
        databarangin.MoveNext
        isitext
    Else
        MsgBox "Sudah data yang paling akhir", vbInformation, "Information"
    End If
End Sub

Private Sub dgbarangin_Click()
    isitext
End Sub

Private Sub Form_Load()
    updateDG
End Sub

Private Sub isitext()
id.Text = dgbarangin.Columns(0).Text
barangid.Text = dgbarangin.Columns(1).Text
userid.Text = dgbarangin.Columns(2).Text
jumlah.Text = dgbarangin.Columns(3).Text
vendor.Text = dgbarangin.Columns(4).Text
End Sub

Private Sub checkEmpty()
If barangid.Text = "" Or userid.Text = "" Or jumlah.Text = "" Or vendor.Text = "" Then
    isEmpty = True
Else
    isEmpty = False
End If
End Sub

Private Sub updateDG()
    Set databarangin = New ADODB.Recordset
    databarangin.Open "select id, id_barang, id_user, jumlah, vendor from barang_masuk", conn
    Set dgbarangin.DataSource = databarangin
End Sub

Private Sub kosongkan()
id.Text = ""
barangid.Text = ""
userid.Text = ""
jumlah.Text = ""
vendor.Text = ""
End Sub
