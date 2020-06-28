Attribute VB_Name = "Module1"
Option Explicit
Public koneksiDB As New ADODB.Connection
Sub bukadb()
    Set koneksiDB = New ADODB.Connection
    koneksiDB.CursorLocation = adUseClient
    koneksiDB.ConnectionString = "DRIVER={MYSQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=db_calonkaryawan;UID=root;Option="
    On Error GoTo pesan
    If koneksiDB.State = adStateClosed Then koneksiDB.Open
        Exit Sub
pesan:
    MsgBox "Maaf ! Tidak Bisa Terkoneksi KeDatabase", vbInformation, "Pesan"
    End
End Sub




