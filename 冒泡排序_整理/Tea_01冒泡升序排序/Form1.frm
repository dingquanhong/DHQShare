VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   13530
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "冒泡（升序排列）"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "交换次数:"
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "比较次数："
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "排序前："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   11055
   End
   Begin VB.Label Label2 
      Caption         =   "排序后："
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   11055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Dim a(1 To 10) As Integer
n = 6
a(1) = 8: a(2) = 6: a(3) = 3: a(4) = 5: a(5) = 9: a(6) = 4
For i = 1 To n     '排序前输出数组各成员
    Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i




For i = 1 To n     '排序后输出数组各成员
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i

End Sub






































