VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   13305
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�Ż�1(�Ѿ�����)for"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��2��(ǰ������)"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�ۺ��Ż���дdo"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�Ż�1(�Ѿ�����)�򻯸�д"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�Ż�1(�Ѿ�����)do"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "��ѭ��������"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "�Ƚϴ�����"
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "����ǰ��"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   11055
   End
   Begin VB.Label Label2 
      Caption         =   "�����"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   11055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click() '��2��(ǰ������)
Dim a(1 To 10) As Integer: Dim flag As Boolean
n = 6
a(1) = 1: a(2) = 2: a(3) = 3: a(4) = 6: a(5) = 5: a(6) = 4
For i = 1 To n     '����ǰ����������Ա
Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
lun = 0: x = 0: Last = 0
For i = 1 To n - 1
    lun = lun + 1 'lun���ڼǴ���ѭ������
    For j = n To i + 1 Step -1
        x = x + 1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
            Last = j
        End If
    Next j
    i = Last - 1
Next i

For i = 1 To n     '���������������Ա
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
Text1.Text = "��ѭ��" + Str(lun) + "��"
Text2.Text = "�Ƚϴ���" + Str(x) + "��"
End Sub



Private Sub Command2_Click() '�Ż�1(�Ѿ�����)for
Dim a(1 To 10) As Integer: Dim flag As Boolean
n = 6
a(1) = 2: a(2) = 3: a(3) = 1: a(4) = 5: a(5) = 6: a(6) = 4
For i = 1 To n     '����ǰ����������Ա
Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
i = 1: lun = 0
For i = 1 To n - 1
    flag = False
    lun = lun + 1 'lun���ڼǴ���ѭ������
    For j = n To i + 1 Step -1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
            flag = True
        End If
    Next j
If flag = False Then Exit For
Next i

For i = 1 To n     '���������������Ա
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
Text1.Text = "��ѭ��" + Str(lun) + "��"
End Sub

Private Sub Command5_Click()  '�Ż�1(�Ѿ�����)do
Dim a(1 To 10) As Integer: Dim flag As Boolean
n = 6
a(1) = 2: a(2) = 3: a(3) = 1: a(4) = 5: a(5) = 6: a(6) = 4
For i = 1 To n     '����ǰ����������Ա
Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
flag = True
i = 1: lun = 0
Do While i <= n - 1 And flag = True
    flag = False
    lun = lun + 1 'lun���ڼǴ���ѭ������
    For j = n To i + 1 Step -1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
            flag = True
        End If
    Next j
    i = i + 1
Loop

For i = 1 To n     '���������������Ա
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
Text1.Text = "��ѭ��" + Str(lun) + "��"
End Sub


Private Sub Command6_Click()  '�Ż�(��ѭ����������)��д��
Dim a(1 To 10) As Integer: Dim flag As Boolean
n = 6
a(1) = 3: a(2) = 1: a(3) = 5: a(4) = 6: a(5) = 2: a(6) = 4
For i = 1 To n     '����ǰ����������Ա
Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
flag = True
lun = 0
Do While flag = True
    flag = False
    lun = lun + 1 'lun���ڼǴ���ѭ������
    For j = n To 2 Step -1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
            flag = True
        End If
    Next j
Loop

For i = 1 To n     '���������������Ա
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
Text1.Text = "��ѭ��" + Str(lun - 1) + "��"
End Sub

Private Sub Command7_Click()  '�ۺ��Ż���дdo
Dim a(1 To 10) As Integer: Dim flag As Boolean
n = 6
a(1) = 1: a(2) = 2: a(3) = 6: a(4) = 3: a(5) = 5: a(6) = 4
For i = 1 To n     '����ǰ����������Ա
Label1.Caption = Label1.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i

flag = True
i = 1
Do While i <= n - 1 And flag = True
    flag = False
    lun = lun + 1 'lun���ڼǴ���ѭ������
    For j = n To i + 1 Step -1
        x = x + 1
        If a(j) < a(j - 1) Then
            t = a(j): a(j) = a(j - 1): a(j - 1) = t
            i = j
            flag = True
        End If
    Next j
Loop

For i = 1 To n     '���������������Ա
Label2.Caption = Label2.Caption + "   a(" + Str(i) + ")=" + Str(a(i))
Next i
Text1.Text = "��ѭ��" + Str(lun) + "��"
Text2.Text = "�Ƚϴ���" + Str(x) + "��"
End Sub



