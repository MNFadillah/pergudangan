Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection

Function koneksi() As Boolean
    Set conn = New ADODB.Connection
    conn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=localhost;Database=db_gudang;User=root;Password=;Option=3;"
    conn.CursorLocation = adUseClient
    koneksi = True
End Function

Sub main()
    If koneksi = True Then
        'frmLogin.Show
        'FrmBarangIn.Show
        'frmBarangOut.Show
        frmUser.Show
    Else
        MsgBox "database gagal, coba sekali lagi"
        End
    End If
End Sub
