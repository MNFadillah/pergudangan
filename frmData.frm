VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmData 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
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
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public table As Integer
Public kolom As String
Dim dataItem As New ADODB.Recordset

Private Sub DataGrid1_Click()
If table = 1 Then
    If kolom = "barang" Then
        FrmBarangIn.barangid = DataGrid1.Columns(0).Text
    Else
        FrmBarangIn.userid = DataGrid1.Columns(0).Text
    End If
ElseIf table = 2 Then
    If kolom = "barang" Then
        'frmBarangOut.barangid = DataGrid1.Columns(0).Text
    Else
        'frmBarangOut.userid = DataGrid1.Columns(0).Text
    End If
End If
Unload Me
End Sub

Private Sub Form_Load()
    Dim query As String
    query = ""
    If kolom = "barang" Then
        query = "select * from barang"
    ElseIf kolom = "user" Then
        query = "select * from user"
    End If
    Set dataItem = New ADODB.Recordset
    dataItem.Open query, conn
    Set DataGrid1.DataSource = dataItem
End Sub
