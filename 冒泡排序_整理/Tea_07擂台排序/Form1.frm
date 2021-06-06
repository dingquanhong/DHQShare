VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "“打擂台”算法"
   ClientHeight    =   3015
   ClientLeft      =   10875
   ClientTop       =   6135
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7230
   Begin VB.CommandButton Command3 
      Caption         =   "“打擂台”算法"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "冒泡（升序排列）"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "排序前："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   11055
   End
   Begin VB.Label Label2 
      Caption         =   "排序后："
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2520
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
jh = 0
bj = 0
a(1) = 8: a(2) = 6: a(3) = 3: a(4) = 5: a(5) = 9: a(6) = 4
For i = 1 To n     '排序前输出数组各成员
    Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i

For i = 1 To n - 1
    For j = n To i + 1 Step -1
 
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
      
        End If
    Next j
Next i

For i = 1 To n     '排序后输出数组各成员
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i


End Sub

Private Sub Command3_Click()   '“打擂台”算法

Dim a(1 To 10) As Integer
n = 6
a(1) = 8: a(2) = 6: a(3) = 3: a(4) = 5: a(5) = 9: a(6) = 4
For i = 1 To n     '排序前输出数组各成员
Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
For i = 1 To n - 1
    For j = i + 1 To n
        If a(j) < a(i) Then
        t = a(j): a(j) = a(i): a(i) = t
        End If
    Next j
Next i
For i = 1 To n     '排序后输出数组各成员
    Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i


End Sub
