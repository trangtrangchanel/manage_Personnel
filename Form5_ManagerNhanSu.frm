VERSION 5.00
Begin VB.Form Form2_NhanSu 
   Caption         =   "Quan Ly"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   LinkTopic       =   "Form2"
   ScaleHeight     =   5115
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVT 
      Height          =   495
      Left            =   4080
      TabIndex        =   21
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdCuoi 
      Caption         =   ">>"
      Height          =   495
      Index           =   3
      Left            =   6480
      TabIndex        =   20
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSau 
      Caption         =   ">"
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   19
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdTruoc 
      Caption         =   "<"
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdDau 
      Caption         =   "<<"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   3360
      Width           =   1575
   End
   Begin VB.OptionButton opnu 
      Caption         =   "Nu"
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.OptionButton opnam 
      Caption         =   "Nam"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtphongban 
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtchucvu 
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtngaysinh 
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtquequan 
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtsdt 
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txttennv 
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtmanv 
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Phong Ban"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Chuc Vu"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Ngay Sinh"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Que Quan"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Gioi Tinh"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "SDT"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ten NV"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Ma NV"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form2_NhanSu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNhanVien As New ADODB.Recordset

Sub getNhanVien()
    If rsNhanVien.State = 1 Then rsNhanVien.Close
    rsNhanVien.Open "NhanSu", cnn, 3, 3
    Set txtmanv.DataSource = rsNhanVien
        txtmanv.DataField = "manv"
    Set txttennv.DataSource = rsNhanVien
        txttennv.DataField = "tennv"
    Set txtngaysinh.DataSource = rsNhanVien
        txtngaysinh.DataField = "ngaysinh"
    Set txtsdt.DataSource = rsNhanVien
        txtsdt.DataField = "sdt"
    Set txtquequan.DataSource = rsNhanVien
        txtquequan.DataField = "quequan"
    Set txtchucvu.DataSource = rsNhanVien
        txtchucvu.DataField = "chucvu"
    Set txtphongban.DataSource = rsNhanVien
        txtphongban.DataField = "phongban"
    Call getGioiTinh
End Sub

Sub getGioiTinh()
    If rsNhanVien!gioitinh Then
        opnam.Value = True
    Else
        opnu.Value = True
    End If
End Sub

Private Sub cmdCuoi_Click(Index As Integer)
    rsNhanVien.MoveLast
    Call getGioiTinh
    txtVT.Text = rsNhanVien.AbsolutePosition
End Sub

Private Sub cmdDau_Click(Index As Integer)
    rsNhanVien.MoveFirst
    Call getGioiTinh
    txtVT.Text = rsNhanVien.AbsolutePosition
End Sub

Private Sub cmdSau_Click(Index As Integer)
    If rsNhanVien.AbsolutePosition < rsNhanVien.RecordCount Then
        rsNhanVien.MoveNext
        Call getGioiTinh
        txtVT.Text = rsNhanVien.AbsolutePosition
    Else
        MsgBox "Ban Ghi Cuoi Cung"
    End If
End Sub

Private Sub cmdTruoc_Click(Index As Integer)
    If rsNhanVien.AbsolutePosition > 1 Then
        rsNhanVien.MovePrevious
        Call getGioiTinh
        txtVT.Text = rsNhanVien.AbsolutePosition
    Else
        MsgBox "Ban Ghi Dau Tien"
    End If
End Sub
Private Sub Form_Load()
    Call getNhanVien
End Sub
