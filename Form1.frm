VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Panel"
   ClientHeight    =   6930
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin VB.Menu mnuDanhMuc 
      Caption         =   "&Danh Muc"
      Begin VB.Menu mnuqlnvien 
         Caption         =   "&Xem Nhan Vien"
      End
      Begin VB.Menu mnuqluong 
         Caption         =   "&Xem Bang Luong"
      End
   End
   Begin VB.Menu mnuquanly 
      Caption         =   "&Quan Ly"
      Begin VB.Menu mnunhanvien 
         Caption         =   "&Nhan Vien"
      End
      Begin VB.Menu mnuqlluong 
         Caption         =   "&Luong"
      End
   End
   Begin VB.Menu mnuthoat 
      Caption         =   "&Thoat"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call ketnoi
    Me.Top = 0: Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Top = (Screen.Height - Me.Height) / 2
If MsgBox("Ban co muon thoat khong?", vbYesNo + vbQuestion + vbSystemModal, "Thoat") = vbNo Then
    Me.Top = 0
    Cancel = 1
Else
    End
End If
End Sub

Private Sub mnunhanvien_Click()
    From4_ThemSUAxoa.Show
End Sub

Private Sub mnuqlnvien_Click()
    Form2_NhanSu.Show
End Sub

Private Sub mnuqluong_Click()
    Form3_Luong.Show
End Sub

Private Sub mnuthoat_Click()
Unload Me
End Sub
