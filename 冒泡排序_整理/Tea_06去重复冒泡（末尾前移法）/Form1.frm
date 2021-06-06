VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
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
   ScaleHeight     =   3720
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "去重复冒泡"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "冒泡排序"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   3975
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
      Height          =   1335
      Left            =   2640
      TabIndex        =   5
      Text            =   "text3"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生n个1-9可重数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1815
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
      Height          =   3120
      Left            =   7080
      TabIndex        =   1
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
      Height          =   3120
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "n:"
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   375
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
      TabIndex        =   3
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
      TabIndex        =   2
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
     b(i) = Int(Rnd * 9 + 1)
     List1.AddItem "b（" + Str(i) + " )=" + Str(b(i))
Next i
End Sub

Private Sub Command2_Click() '冒泡
For i = 1 To n
a(i) = b(i)
Next i


For i = 1 To n - 1
   
    For j = n To i + 1 Step -1
       
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
        End If
    Next j
Next i

List2.Clear
For i = 1 To n
    List2.AddItem "a（" + Str(i) + " )=" + Str(a(i))
Next i

End Sub

Private Sub Command3_Click()   '去重复冒泡
For i = 1 To n
a(i) = b(i)
Next i

bottom = n
For i = 1 To bottom - 1
    For j = bottom To i + 1 Step -1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
        ElseIf a(j) = a(j - 1) Then
            a(j) = a(bottom)
            bottom = bottom - 1
        End If
    Next j
Next i
List2.Clear
For i = 1 To bottom
    List2.AddItem "a（" + Str(i) + " )=" + Str(a(i))
Next i

End Sub


