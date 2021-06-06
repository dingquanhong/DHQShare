VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "优化（综合优化）"
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "冒泡排序"
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Text            =   "20"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生n个1-99不重复数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   7080
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "n:"
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "比较次数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "外循环轮数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "排序后："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "排序前："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim a(1 To 100) As Integer
Dim b(1 To 100) As Integer


Private Sub Command1_Click()
List1.Clear
n = Val(Text3.Text)
For i = 1 To n
     b(i) = Int(Rnd * 99 + 1)
     flag = True
     For j = 1 To i - 1
        If b(i) = b(j) Then
            i = i - 1
            flag = False
            Exit For
           
        End If
     Next j
     If flag = True Then List1.AddItem "b（" + Str(i) + " )=" + Str(b(i))
Next i
End Sub

Private Sub Command2_Click()
For i = 1 To n
a(i) = b(i)
Next i

lun = 0
For i = 1 To n - 1
    lun = lun + 1 'lun用于记录外循环几轮
    For j = n To 2 Step -1
        bj = bj + 1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
        End If
    Next j
Next i

List2.Clear
For i = 1 To n
    List2.AddItem "a（" + Str(i) + " )=" + Str(a(i))
Next i
Text1.Text = Str(lun)
Text2.Text = Str(bj)
End Sub

Private Sub Command3_Click()
For i = 1 To n
a(i) = b(i)
Next i

flag = True
i = 1: lun = 0
Do While flag
    flag = False
    lun = lun + 1 'lun用于记录外循环几轮
    For j = n To i + 1 Step -1
        bj = bj + 1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
            i = j
            flag = True
        End If
    Next j
Loop

List2.Clear
For i = 1 To n
    List2.AddItem "a（" + Str(i) + " )=" + Str(a(i))
Next i
Text1.Text = Str(lun)
Text2.Text = Str(bj)
End Sub


