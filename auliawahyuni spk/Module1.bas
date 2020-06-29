Attribute VB_Name = "Module1"
Option Explicit
Public db_pelangganterbaik As New ADODB.Connection
Sub bukadb()
    Set db_pelangganterbaik = New ADODB.Connection
    db_pelangganterbaik.CursorLocation = adUseClient
    db_pelangganterbaik.ConnectionString = "driver={mysql odbc 3.51 driver};server=localhost;database=db_pelanggnaterbaik;uid=root;option"
    On Error GoTo pesan
    If db_pelangganterbaik.State = adStateClosed Then db_pelangganterbaik.Open
Exit Sub
pesan:
 MsgBox "Maaf ! Tidak Bisa Terkoneksi KeDatabase", vbInformation, "Pesan"
    End
End Sub

