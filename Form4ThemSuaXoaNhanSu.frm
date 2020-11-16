VERSION 5.00
Begin VB.Form From4_ThemSUAxoa 
   Caption         =   "Form2"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10515
   LinkTopic       =   "Form2"
   ScaleHeight     =   5610
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   735
      Left            =   9360
      TabIndex        =   26
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "OK"
      Height          =   735
      Left            =   8400
      TabIndex        =   25
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdXoa 
      Caption         =   "Xoa"
      Height          =   615
      Left            =   8400
      TabIndex        =   24
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSua 
      Caption         =   "Sua"
      Height          =   495
      Left            =   8400
      TabIndex        =   23
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdthem 
      Caption         =   "Them"
      Height          =   435
      Left            =   8400
      TabIndex        =   22
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtmanv 
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txttennv 
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtsdt 
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtquequan 
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtngaysinh 
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtchucvu 
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtphongban 
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton opnam 
      Caption         =   "Nam"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton opnu 
      Caption         =   "Nu"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdDau 
      Caption         =   "<<"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdTruoc 
      Caption         =   "<"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSau 
      Caption         =   ">"
      Height          =   495
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCuoi 
      Caption         =   ">>"
      Height          =   495
      Index           =   3
      Left            =   6360
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtVT 
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Ma NV"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ten NV"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "SDT"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Gioi Tinh"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Que Quan"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Ngay Sinh"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Chuc Vu"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Phong Ban"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "From4_ThemSUAxoa"
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

Private Sub cmdNo_Click()
    rsNhanVien.CancelUpdate
    Call SangMo(True)
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

Private Sub SangMo(b As Boolean)
    txtmanv.Locked = b
    txttennv.Locked = b
    txtsdt.Locked = b
    txtquequan.Locked = b
    txtngaysinh.Locked = b
    txtchucvu.Locked = b
    txtphongban.Locked = b
    
    cmdthem.Enabled = b
    cmdSua.Enabled = b
    cmdXoa.Enabled = b
    'cmdDau.Enabled = b'
    'cmdTruoc.Enabled = b'
    'cmdSau.Enabled = b'
    'cmdCuoi.Enabled = b'
    
    cmdYes.Enabled = Not b
    cmdNo.Enabled = Not b
End Sub

Private Sub cmdSua_Click()
    Call SangMo(False)
    txtmanv.SetFocus
End Sub

Private Sub cmdthem_Click()
    rsNhanVien.AddNew
    Call SangMo(False)
    txtmanv.SetFocus
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

Private Sub cmdXoa_Click()
    Dim iVT As Long
    If MsgBox("Xoa Khong?", vbYesNo + vbQuestion, "Xoa") = vbYes Then
        iVT = rsNhanVien.AbsolutePosition
        rsNhanVien.Delete
        If Not rsNhanVien.EOF Then
            If iVT < rsNhanVien.RecordCount Then
                rsNhanVien.AbsolutePosition = iVT
            Else
                rsNhanVien.AbsolutePosition = iVT - 1
            End If
        End If
    End If
End Sub

Private Sub cmdYes_Click()
    rsNhanVien!gioitinh = IIf(opnam.Value, -1, 0)
    rsNhanVien.Update
    Call SangMo(True)
End Sub

Private Sub Form_Load()
    Call getNhanVien
End Sub

