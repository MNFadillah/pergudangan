VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      MaskColor       =   &H00808000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Login Aplikasi Pergudangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isempty As Boolean

Private Sub Command1_Click()
    Call checkEmpty
    If isempty = False Then
        Dim login As New ADODB.Recordset
        login.Open "select*from user where username = '" & Text1.Text & "' and password = md5('" & Text2.Text & "')", conn
        If login.RecordCount > 0 Then
            MsgBox "Anda Berhasil Login", vbInformation, "Login Berhasil"
            frmMain.Show
            Unload frmLogin
        Else
            MsgBox "Username/Password salah", vbCritical, "Login Gagal"
            kosongkan
        End If
    End If
    
End Sub

Private Sub checkEmpty()
    If Text1.Text = "" Or Text2.Text = "" Then
        isempty = True
    Else
        isempty = False
    End If
    
End Sub

Private Sub kosongkan()
Text1.Text = ""
Text2.Text = ""
End Sub
