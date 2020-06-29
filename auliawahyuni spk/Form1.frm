VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   18
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "keluar"
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7200
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=db_pelangganterbaik"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "db_pelangganterbaik"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbl_pelanggan"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "edit"
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "batal"
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hapus"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton sim 
      Caption         =   "simpan"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2295
      Left            =   2280
      TabIndex        =   7
      Top             =   5880
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
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
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "CARI"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "pekerjaan"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "alamat"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "telepon"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "tempat lahir"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "nama pelanggan"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ID pelanggan"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA PELANGGAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db_pelangganterbaik As New ADODB.Recordset

Private Sub Command3_Click()
Call kosong
End Sub

Private Sub Command4_Click()
konekdb_pelangganterbaik.Execute "update tbl_pelanggan set nm_pelanggan ='" & Text2.Text & "',tempat_lahir='" & Text3.Text & "',tlp='" & Text4.Text & "',alamat='" & Text5.Text & "',pekerjaan='" & Text5.Text & "' where id_pelanggan='" & Text1.Text & "'"
MsgBox "data berhasil di edit", vbInformation, "pesan"
Call segar
Call kosong
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Dim hapus As String
hapus = MsgBox("yakin akan menghapus data ini", vbYesNo, "pesan")
If hapus = vbYes Then
konekdb_pelangganterbaik.Execute "delete from tbl_pelanggan where id_pelanggan='" & Text1.Text & "'"
Call segar
Call kosong
Text1.SetFocus
MsgBox "data telah dihapus", vbExclamation, "pesan"
End If
End Sub

Private Sub Command1_Click()
    If Text1.Text = "" Then
    MsgBox "id pelanggan kosong", vbExclamation, "pesan"
    Text1.SetFocus
    Exit Sub
    End If
    If Text2.Text = "" Then
    MsgBox "nama pelanggan kosong", vbExclamation, "pesan"
    Text2.SetFocus
    Exit Sub
    End If
    If Text3.Text = "" Then
    MsgBox "tempat lahir  kosong", vbExclamation, "pesan"
    Text3.SetFocus
    Exit Sub
    End If
    If Text4.Text = "" Then
    MsgBox "telepon  kosong", vbExclamation, "pesan"
    Text4.SetFocus
    Exit Sub
    End If
    If Text5.Text = "" Then
    MsgBox "alamat  kosong", vbExclamation, "pesan"
    Text5.SetFocus
    Exit Sub
    End If
    If Text6.Text = "" Then
    MsgBox "pekerjaan  kosong", vbExclamation, "pesan"
    Text6.SetFocus
    Exit Sub
    End If
Set db_pelangganterbaik = New ADODB.Recordset
db_pelangganterbaik.Open "select * from tbl_pelanggan where id_pelanggan='" & Text1.Text & "'", konekdb_pelangganterbaik
If Not db_pelangganterbaik.EOF Then
MsgBox "id pelanggan sudah digunakan", vbCritical, "pesan"
Text1.Text = ""
Text1.SetFocus
Exit Sub
Else
konekdb_pelangganterbaik.Execute "insert into tbl_pelanggan(id_pelanggan,nm_pelanggan,tempat_lahir,tlp,alamat,pekerjaan) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "')"
MsgBox "DATA TERSIMPAN"
Call segar
Text1.SetFocus
End If

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = db_pelangganterbaik!id_pelanggan
Text2.Text = db_pelangganterbaik!nm_pelanggan
Text3.Text = db_pelangganterbaik!tempat_lahir
Text4.Text = db_pelangganterbaik!tlp
Text5.Text = db_pelangganterbaik!alamat
Text6.Text = db_pelangganterbaik!pekerjaan

End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = db_pelangganterbaik
With db_pelangganterbaik
End With
Call edit_grid

Combo1.AddItem
End Sub
Sub tampil_data()
Set db_pelangganterbaik = New ADODB.Recordset
db_pelangganterbaik.ActiveConnection = konekdb_pelangganterbaik
db_pelangganterbaik.CursorLocation = adUseClient
db_pelangganterbaik.LockType = adLockOptimistic
db_pelangganterbaik.Source = "select * from tbl_pelanggan"
db_pelangganterbaik.Open
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "id pelanggan"
    .Columns(1).Caption = "Nama pelanggan"
    .Columns(2).Caption = "tempat lahir"
    .Columns(3).Caption = "telepon"
    .Columns(4).Caption = "alamat"
    .Columns(5).Caption = "pekerjaan"
    .Columns(0).Width = 2000
    .Columns(1).Width = 2000
    .Columns(2).Width = 3000
End With
End Sub
Sub segar()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = db_pelangganterbaik
With DataGrid1
Call edit_grid
End With
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub


Private Sub text4_Change()
Set db_pelangganterbaik = New ADODB.Recordset
db_pelangganterbaik.Open "select * from tbl_lokasi where id_pelanggan like '%" & Text7.Text & "%'", konekdb_pelangganterbaik
If Not db_pelangganterbaik.EOF Then
Set DataGrid1.DataSource = db_pelangganterbaik
Call edit_grid
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If db_pelangganterbaik.State = adStateOpen Then db_pelangganterbaik.Close
db_pelangganterbaik.Open "select * from tbl_pelanggan where id_pelanggan like '%" & Text7.Text & "%'", konekdb_pelangganterbaik
If Not db_pelangganterbaik.EOF Then
Text1.Text = db_pelangganterbaik!id_pelanggan
Text2.Text = db_pelangganterbaik!nm_pelanggan
Text3.Text = db_pelangganterbaik!tempat_lahir
Text4.Text = db_pelangganterbaik!tlp
Text5.Text = db_pelangganterbaik!alamat
Text6.Text = db_pelangganterbaik!pekerjaan

Call segar
End If
End If
End Sub


