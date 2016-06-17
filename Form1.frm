VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "高精度加减乘除阶乘模块"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   3375
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5160
      Width           =   10695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "高精度阶乘"
      Height          =   735
      Left            =   9360
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "高精度除法"
      Height          =   735
      Left            =   7200
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "高精度乘法"
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "高精度减法"
      Height          =   735
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "高精度加法"
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "1000"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   $"Form1.frx":0000
      Top             =   1680
      Width           =   10335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   $"Form1.frx":00BD
      Top             =   720
      Width           =   10335
   End
   Begin VB.Label Label4 
      Caption         =   "高精度除法和高精度阶乘只能以长整型举例，不然会溢出"
      Height          =   615
      Left            =   7440
      TabIndex        =   11
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "除法保留小数位"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "数B"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "数A"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text4.Text = Jia(Text1.Text, Text2.Text)
End Sub

Private Sub Command2_Click()
Text4.Text = Jian(Text1.Text, Text2.Text)
End Sub

Private Sub Command3_Click()
Text4.Text = Cheng(Text1.Text, Text2.Text)
End Sub

Private Sub Command4_Click()
Text1.Text = 17
Text2.Text = 7
Text4.Text = Chu(Text1.Text, Text2.Text, Val(Text3.Text))
End Sub

Private Sub Command5_Click()
If Val(Text1.Text) > 2147483647 Then Text1.Text = "2147483647"
Text4.Text = JieCheng(Text1.Text)
End Sub
