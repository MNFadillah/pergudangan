VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUser 
   Caption         =   "User"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "DATA"
      Height          =   4935
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command7 
         Caption         =   ">>"
         Height          =   615
         Left            =   4200
         TabIndex        =   19
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   ">"
         Height          =   615
         Left            =   3360
         TabIndex        =   18
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<"
         Height          =   615
         Left            =   2520
         TabIndex        =   17
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<<"
         Height          =   615
         Left            =   1680
         TabIndex        =   16
         Top             =   4200
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid dgbarang 
         Height          =   3735
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6588
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   2280
         TabIndex        =   14
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   1200
         TabIndex        =   13
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox password 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox username 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox ktp 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox nama 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox id 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "USERNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "KTP"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "NAMA"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim databarang As New ADODB.Recordset
Dim isempty As Boolean

Private Sub Command1_Click()
checkEmpty
If isempty = False Then
    On Error GoTo addError
    query = "insert into user values ('', '" & nama.Text & "', '" & ktp.Text & "', '" & username.Text & "', md5('" & password.Text & "'))"
    conn.Execute query
    MsgBox "Berhasil menambahkan data", vbInformation, "Sukses"
    updateDG
    kosongkan
    Exit Sub
addError:
    MsgBox "Gagal menambahkan data " & Err.Description, vbInformation
Else
    MsgBox "Pilih data terlebih dahulu", vbCritical, "Warning"
End If
End Sub

Private Sub Command2_Click()
checkEmpty
If isempty = False Then
    On Error GoTo addError
    query = "update user set nama = '" & nama.Text & "', ktp = '" & ktp.Text & "', username = '" & username.Text & "', password= md5('" & password.Text & "') where id = " & id.Text & ""
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
If isempty = False Then
    On Error GoTo addError
    query = "delete from user where id = " & id.Text & ""
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

Private Sub Command4_Click()
If Not databarang.AbsolutePosition = 0 Then
        databarang.MoveFirst
        isitext
    Else
        'MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub Command5_Click()
If Not databarang.AbsolutePosition = 1 Then
        databarang.MovePrevious
        isitext
    Else
        MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub Command6_Click()
If Not databarang.AbsolutePosition = databarang.RecordCount Then
        databarang.MoveNext
        isitext
    Else
        MsgBox "Sudah data yang paling akhir", vbInformation, "Information"
    End If
End Sub

Private Sub Command7_Click()
If Not databarang.AbsolutePosition = databarang.RecordCount Then
        databarang.MoveLast
        isitext
    Else
        'MsgBox "Sudah data yang paling awal", vbInformation, "Information"
    End If
End Sub

Private Sub dgbarang_Click()
isitext
End Sub

Private Sub Form_Load()
    updateDG
End Sub

Private Sub isitext()
id.Text = dgBarang.Columns(0).Text
nama.Text = dgBarang.Columns(1).Text
ktp.Text = dgBarang.Columns(2).Text
username.Text = dgBarang.Columns(3).Text
password.Text = dgBarang.Columns(4).Text
End Sub

Private Sub checkEmpty()
If nama.Text = "" Or ktp.Text = "" Or username.Text = "" Or password.Text = "" Then
    isempty = True
Else
    isempty = False
End If
End Sub

Private Sub updateDG()
    Set databarang = New ADODB.Recordset
    databarang.Open "select * from user", conn
    Set dgBarang.DataSource = databarang
End Sub

Private Sub kosongkan()
id.Text = ""
nama.Text = ""
ktp.Text = ""
username.Text = ""
password.Text = ""
End Sub

