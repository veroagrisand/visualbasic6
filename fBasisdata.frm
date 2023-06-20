VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fBasisdata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basis data"
   ClientHeight    =   270
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10200
      Top             =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Proses Data"
      Height          =   2175
      Left            =   360
      TabIndex        =   2
      Top             =   6720
      Width           =   11655
      Begin VB.CommandButton Command3 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   8880
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit/Update"
         Height          =   495
         Left            =   4560
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview Data"
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "fBasisdata.frx":0000
         Height          =   1935
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\PBO VB6\dB-TISem4A.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Enabled         =   0   'False
         Exclusive       =   0   'False
         Height          =   465
         Left            =   7200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "tMHS"
         Top             =   960
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   1920
         TabIndex        =   9
         Text            =   "---Kode Jurusan--"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   8
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   7
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   7440
         TabIndex        =   14
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Jurusan"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Nama"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nomor BP"
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Aplikasi Basis Data Sederhana"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   5295
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu smExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "fBasisdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub clear()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
End Sub
Private Sub Command1_Click()
tanya = MsgBox("yakin data ini ingin didimpaan [y/t]?", vbYesNo, "simpan data")
If tanya = vbYes Then
    'proses simpan data'
    Data1.Recordset.AddNew
    Data1.Recordset(0) = Text1.Text
     Data1.Recordset(1) = Text2.Text
      Data1.Recordset(2) = Combo1.Text
       Data1.Recordset.Update
       Data1.Refresh
       MsgBox "data sukses disimpan", vbInformation, ""
       Text1.Text = "": Text2.Text = "": Combo1.Text = ""
       Command1.Enabled = False
      Command2.Enabled = False
      Command3.Enabled = False
      Text1.SetFocus
      End If
End Sub

Private Sub Command2_Click()
tanya = MsgBox("yakin data ini ingin di edit [y/t]?", vbYesNo, "Edit  data")
If tanya = vbYes Then
    'proses simpan data'
    Data1.Recordset.Edit
     Data1.Recordset(1) = Text2.Text
      Data1.Recordset(2) = Combo1.Text
       Data1.Recordset.Update
       Data1.UpdateRecord
       Data1.Refresh
       MsgBox "data sukses diupdate", vbInformation, ""
       Text1.Text = "": Text2.Text = "": Combo1.Text = ""
       Command1.Enabled = False
      Command2.Enabled = False
      Command3.Enabled = False
      Text1.SetFocus
      End If
End Sub

Private Sub Command3_Click()
tanya = MsgBox("yakin data ini ingin dihapus [y/t]?", vbYesNo, "Hapus data")
If tanya = vbYes Then
    'proses simpan data'
       Data1.Recordset.Delete
       Data1.Recordset.MoveNext
       Data1.Refresh
       MsgBox "data sukses dihapus", vbInformation, ""
       Text1.Text = "": Text2.Text = "": Combo1.Text = ""
       Command1.Enabled = False
      Command2.Enabled = False
      Command3.Enabled = False
      Text1.SetFocus
      End If
End Sub

Private Sub smExit_Click()
Timer1.Enabled = True
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(Text1.Text) > 9 Then
    'Label5.Caption = "Sedang Mencari"
    Data1.Recordset.Index = "nomorBP"
    Data1.Recordset.Seek "=", Text1.Text
    If Data1.Recordset.NoMatch Then
        MsgBox "Data " & Text1.Text & " TIdak Ditemukan, dianggap Data baru. Silahkan melengkapi", vbInformation, "Pencarian Data"
        'Text1.Text = "": Combo1.Text
        Text2.SetFocus
        Command1.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = False
        
        Else
        Text2.Text = Data1.Recordset(1)
        Combo1.Text = Data1.Recordset(2)
        Command1.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = True
    End If
    Else
    
    Label5.Caption = "..."
    
    End If

End Sub

Private Sub Timer1_Timer()
If fBasisdata.Height > 720 Then
fBasisdata.Height = fBasisdata.Height - 35
fBasisdata.Top = fBasisdata.Top + 20

Else
fUTAMA.Enabled = True
fBasisdata.Visible = False
fBasisdata.Enabled = False
End If
End Sub
