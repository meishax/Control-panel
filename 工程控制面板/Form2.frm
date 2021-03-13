VERSION 5.00
Begin VB.Form Form2 
   Caption         =   $"Form2.frx":0000
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12915
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   12915
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "修改工况三"
      Height          =   1095
      Left            =   5280
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "修改工况二"
      Height          =   1095
      Left            =   3360
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   1095
      Left            =   7200
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改工况一"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Form1.Show
Form2.Enabled = False
Form2.Hide
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Command4_Click()
Form3.Enabled = True
Form3.Show
Form2.Enabled = False
Form2.Hide
End Sub
Private Sub Command5_Click()
Form4.Enabled = True
Form4.Show
Form2.Enabled = False
Form2.Hide
End Sub
Private Sub Form_Load()
Form1.Enabled = False
Form1.Hide
Form3.Enabled = False
Form3.Hide
Form4.Enabled = False
Form4.Hide
End Sub
