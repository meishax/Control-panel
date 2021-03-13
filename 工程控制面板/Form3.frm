VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12285
   LinkTopic       =   "Form3"
   ScaleHeight     =   8460
   ScaleWidth      =   12285
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Text            =   "请输入内容"
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改参数"
      Height          =   1215
      Left            =   1200
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   1215
      Left            =   4320
      TabIndex        =   5
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Text            =   "请输入内容"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Text            =   "请输入内容"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Text            =   "请输入内容"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Text            =   "请输入内容"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "查看结果"
      Height          =   1215
      Left            =   7440
      TabIndex        =   0
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "前屋面腹杆排列密度"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "后屋面腹杆排列密度"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "后屋面角"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "温室半径"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "垂直距离"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data_dir As String
Dim SD As String     'SD
Dim XD As String    'XD
Dim DEG_BACK As String   'DEG_BACK
Dim R_FRONT As String  'R_Front
Dim D_D_HIGH As String  'D_D_HIGH
Dim data_str As String     'String
Dim data_strShow As String
Dim midStr0 As String
Dim midStr1 As String
Dim midStr2 As String
Dim midStr3 As String
Dim midStr4 As String
Dim txtchangeF As Boolean '
Private Sub Command1_Click()
If txtchangeF Then
    data_dir = App.Path & "\operating condition1.txt"
    
    midStr0 = "SD=" & Text1.Text
    Call ChangeLine(data_dir, 6, midStr0)
  
    midStr1 = "XD=" & Text2.Text
    Call ChangeLine(data_dir, 7, midStr1)
    
    midStr2 = "$DEG_BACK=" & Text3.Text
    Call ChangeLine(data_dir, 8, midStr2)
   
    midStr2 = "R_FRONT=" & Text4.Text
    Call ChangeLine(data_dir, 9, midStr2)
   
    midStr2 = "$D_D_HIGH=" & Text5.Text
    Call ChangeLine(data_dir, 10, midStr2)
    
    MsgBox "参数修改成功！"
    txtchangeF = False
End If
End Sub
Private Sub Command2_Click()
Form2.Show
Form2.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()
Dim SD_VariantShow
Dim XD_VariantShow
Dim DEG_BACK_VariantShow
Dim R_FRONT_VariantShow
Dim D_D_HIGH_VariantShow
Dim SD_StringShow As String
Dim XD_StringShow As String
Dim R_FRONT_StringShow As String
Dim DEG_BACK_StringShow As String
Dim D_D_HIGH_StringShow As String
Dim LShow As Integer


Dim j As Integer
j = 0
Open App.Path & "\operating condition2show.txt" For Input As #9
    Do While Not EOF(9)
        Line Input #9, data_strShow
        j = j + 1
        Select Case j
        
        Case 1
        
        SD_VariantShow = Split(Trim(data_strShow))
        SD_StringShow = Trim(CStr(SD_VariantShow(0)))
        
        Case 2
        XD_VariantShow = Split(Trim(data_strShow))
        XD_StringShow = Trim(CStr(XD_VariantShow(0)))
        
        Case 3
        DEG_BACK_VariantShow = Split(Trim(data_strShow))
        DEG_BACK_StringShow = Trim(CStr(DEG_BACK_VariantShow(0)))

        Case 4
        R_FRONT_VariantShow = Split(Trim(data_strShow))
        R_FRONT_StringShow = Trim(CStr(R_FRONT_VariantShow(0)))
        
        Case 5
        D_D_HIGH_VariantShow = Split(Trim(data_strShow))
        D_D_HIGH_StringShow = Trim(CStr(D_D_HIGH_VariantShow(0)))
        
        End Select
    Loop
    MsgBox (SD_StringShow & Chr(13) & XD_StringShow & Chr(13) & DEG_BACK_StringShow & Chr(13) & R_FRONT_StringShow & Chr(13) & D_D_HIGH_StringShow)
Close #9
End Sub


Private Sub Form_Load()
Dim SD_Variant
Dim XD_Variant
Dim DEG_BACK_Variant
Dim R_FRONT_Variant
Dim D_D_HIGH_Variant
Dim SD_String As String
Dim XD_String As String
Dim R_FRONT_String As String
Dim DEG_BACK_String As String
Dim D_D_HIGH_String As String
Dim L As Integer


Dim j As Integer
j = 0
Open App.Path & "\operating condition2.txt" For Input As #2
    Do While Not EOF(2)
        Line Input #2, data_str
        j = j + 1
        Select Case j
        
        Case 6
        
        SD_Variant = Split(Trim(data_str))
        SD_String = Trim(CStr(SD_Variant(0)))
        L = Len(SD_String)
        Text1.Text = Mid(SD_String, 4, L)
        
        Case 7
        XD_Variant = Split(Trim(data_str))
        XD_String = Trim(CStr(XD_Variant(0)))
        L = Len(XD_String)
        Text2.Text = Mid(XD_String, 4, L)
        
        Case 8
        DEG_BACK_Variant = Split(Trim(data_str))
        DEG_BACK_String = Trim(CStr(DEG_BACK_Variant(0)))
        L = Len(DEG_BACK_String)
        Text3.Text = Mid(DEG_BACK_String, 10, L)
        
        Case 9
        R_FRONT_Variant = Split(Trim(data_str))
        R_FRONT_String = Trim(CStr(R_FRONT_Variant(0)))
        L = Len(R_FRONT_String)
        Text4.Text = Mid(R_FRONT_String, 9, L)
        
        Case 10
        D_D_HIGH_Variant = Split(Trim(data_str))
        D_D_HIGH_String = Trim(CStr(D_D_HIGH_Variant(0)))
        L = Len(D_D_HIGH_String)
        Text5.Text = Mid(D_D_HIGH_String, 10, L)
        
        End Select
    Loop
Close #2
txtchangeF = False
End Sub
Private Sub Text1_Change()
txtchangeF = True
End Sub
Private Function ChangeLine(strFile As String, RLine As Long, NewStr As String)
Dim s As String, n As String, i As Long
i = 1
'//打开源文件
Open strFile For Input As #1
Do Until EOF(1)
    Line Input #1, s
    If RLine = i Then '如果是指定的行数就进行下面的操作
            s = NewStr
            n = n & s & vbCrLf '将空字符串赋给变量n,以保持源文件的行数
        Else    '如果不是指定的行数,就将s的内容赋给变量n 以存储数据
        n = n & s & vbCrLf   '将s的内容赋给n 并以一个回车符号结束....
    End If
    i = i + 1
Loop
Close #1

   '//写入新文件,如果和源文件同名则会覆盖源文件
Open strFile For Output As #2
Print #2, n '将n变量里的数据写入新文件
Close #2
End Function

Private Sub Text2_Change()
txtchangeF = True
End Sub

Private Sub Text3_Change()
txtchangeF = True
End Sub

Private Sub Text4_Change()
txtchangeF = True
End Sub

Private Sub Text5_Change()
txtchangeF = True
End Sub

