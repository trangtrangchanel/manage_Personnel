VERSION 5.00
Begin VB.Form Form3_Luong 
   Caption         =   "Form2"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVT 
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdCuoi 
      Caption         =   ">>"
      Height          =   615
      Left            =   5160
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSau 
      Caption         =   ">"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdTruoc 
      Caption         =   "<"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdDau 
      Caption         =   "<<"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtthucnhan 
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtthangnhanluong 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtbaohiem 
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtthue 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtluong 
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtmanv 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Thuc Nhan"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Thue"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Thanh Nhan Luong"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Bao Hiem"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Luong"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Ma NV"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form3_Luong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLuong As New ADODB.Recordset
Dim getluongs, getbaohiem, getthue, thucnhan As Double
Sub getluong()
    If rsLuong.State = 1 Then rsLuong.Close
    rsLuong.Open "Luong", cnn, 3, 3
    Set txtmanv.DataSource = rsLuong
        txtmanv.DataField = "manv"
    Set txtluong.DataSource = rsLuong
        txtluong.DataField = "luong"
    Set txtthue.DataSource = rsLuong
        txtthue.DataField = "thue"
    Set txtbaohiem.DataSource = rsLuong
        txtbaohiem.DataField = "baohiem"
    Set txtthangnhanluong.DataSource = rsLuong
        txtthangnhanluong.DataField = "thang"

    getluongs = rsLuong!luong
    getbaohiem = rsLuong!baohiem
    getthue = ((getluongs * rsLuong!thue) / 100)
    thucnhan = getluongs - getthue + getbaohiem
    txtthucnhan.Text = thucnhan
    
End Sub

Private Sub cmdCuoi_Click()
    rsLuong.MoveLast
End Sub

Private Sub cmdDau_Click()
    rsLuong.MoveFirst
End Sub

Private Sub cmdSau_Click()
    If rsLuong.AbsolutePosition < rsLuong.RecordCount Then
        rsLuong.MoveNext
    Else
        MsgBox "Day la Ban Ghi Cuoi Cung"
    End If
End Sub

Private Sub cmdTruoc_Click()
    If rsLuong.AbsolutePosition > 1 Then
        rsLuong.MovePrevious
    Else
        MsgBox "Day la Ban Ghi Dau Tien"
    End If
End Sub

Private Sub Form_Load()
    Call getluong
End Sub

Private Sub txtmanv_Change()
    txtVT = rsLuong.AbsolutePosition & "/" & rsLuong.RecordCount
End Sub

