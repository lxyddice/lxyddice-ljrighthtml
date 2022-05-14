VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "简单的html生成器"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   11610
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command11 
      Caption         =   "退出"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "输出"
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "加尾"
      Height          =   495
      Left            =   10440
      TabIndex        =   35
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   9480
      TabIndex        =   33
      Text            =   "未启用"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "加入<a>"
      Height          =   255
      Left            =   10320
      TabIndex        =   31
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   8160
      TabIndex        =   30
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   8160
      TabIndex        =   29
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   8520
      TabIndex        =   26
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   8520
      TabIndex        =   24
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   8280
      TabIndex        =   22
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "加入<a>"
      Height          =   375
      Left            =   10560
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "加粗（x）"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   8880
      TabIndex        =   14
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   8880
      TabIndex        =   13
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "加入<a>"
      Height          =   375
      Left            =   10560
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "点击后删除文本框（√）"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "加入<p>"
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   7800
      TabIndex        =   7
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "加入"
      Height          =   495
      Left            =   10920
      TabIndex        =   6
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   8160
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加头"
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   6240
      Width           =   9495
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label Label15 
      Caption         =   "请在输出txt后把*/*/改为“""”"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label14 
      Caption         =   "大小："
      Height          =   255
      Left            =   9000
      TabIndex        =   32
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "宽："
      Height          =   375
      Left            =   7800
      TabIndex        =   28
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "长："
      Height          =   375
      Left            =   7800
      TabIndex        =   27
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "超链接："
      Height          =   255
      Left            =   7800
      TabIndex        =   25
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "文件名："
      Height          =   255
      Left            =   7800
      TabIndex        =   23
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "路径："
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "插入图片"
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "-------------------------------------------"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "-------------------------------------------"
      Height          =   255
      Left            =   7680
      TabIndex        =   18
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "超链接名字："
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "超链接地址："
      Height          =   255
      Left            =   7800
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "-------------------------------------------"
      Height          =   135
      Left            =   7680
      TabIndex        =   11
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "最终编辑框："
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "标题："
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jc As Long
Dim tg As Long
Dim sc As Long
Dim xh As Long
Private Sub Command1_Click()
List1.AddItem "<!DOCTYPE html>"
List1.AddItem "<html>"
List1.AddItem "<head>"
List1.AddItem "<meta charset=" & "*/*/utf-8*/*/" & ">"
List1.AddItem "<title>" & Text2.Text & "</title>"
List1.AddItem "</head>"
List1.AddItem "<body>"
End Sub

Private Sub Command10_Click()
For i = 0 To List1.ListCount - 1
Open App.Path & "\" & xh & ".txt" For Append As #1
Print #1, List1.List(i)
Close #1
Next
Close fn
End Sub

Private Sub Command11_Click()
End
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "不能输入空字符！"
Else
List1.AddItem Text1.Text
End If
If sc = 0 Then
Text1.Text = ""
End If
End Sub

Private Sub Command3_Click()
Text1.Text = "<p>" & Text3.Text & "</p>"
If sc = 0 Then
Text3.Text = ""
End If
End Sub

Private Sub Command4_Click()
If sc = 1 Then
sc = 0
Command4.Caption = "点击后删除文本框（√）"
Else
sc = 1
Command4.Caption = "点击后删除文本框（x）"
End If


End Sub

Private Sub Command5_Click()
Text1.Text = "<a>" & Text3.Text & "</a>"
If sc = 0 Then
Text3.Text = ""
End If
End Sub

Private Sub Command6_Click()
If jc = 1 Then
jc = 0
Command6.Caption = "加粗（x）"
Else
jc = 1
Command6.Caption = "加粗（√）"
End If

End Sub

Private Sub Command7_Click()
If Text4.Text = "" Then
MsgBox "不能为空！"
Else
If Text5.Text = "" Then
MsgBox "不能为空"
Else
If jc = 0 Then
Text1.Text = "<a href=*/*/" & Text4.Text & "*/*/>" & Text5.Text & "</a>"
Else
Text1.Text = "<strong>" & "<a href=*/*/" & Text4.Text & "*/*/>" & Text5.Text & "</a>" & "</strong>"
End If
End If
End If
If sc = 0 Then
Text4.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub Command8_Click()
List1.AddItem "</body>"
List1.AddItem "</html>"
End Sub

Private Sub Command9_Click()
tg = 1

If Text9.Text = "" Then
MsgBox "长不能为空"
tg = 0
End If

If Text10.Text = "" Then
MsgBox "宽不能为空"
tg = 0
End If

If Text7.Text = "" Then
MsgBox "文件名不能为空"
tg = 0
End If

If tg = 1 Then
If Text8.Text = "" Then

Text1.Text = "<img src =*/*/" & "/" & Text6.Text & "/" & Text7.Text & "*/*/ " & "alt=*/*/Pulpit rock*/*/ weith=*/*/" & Text9.Text & "*/*/ height=*/*/" & Text10.Text & "*/*/>"

Else

List1.AddItem "<a href=*/*/" & Text8.Text & "*/*/>"

Text1.Text = "<img border=*/*/0*/*/ src=*/*/" & "/" & Text6.Text & "/" & Text7.Text & "*/*/ " & "alt=*/*/Pulpit rock*/*/ weith=*/*/" & Text9.Text & "*/*/ height=*/*/" & Text10.Text & "*/*/>"


End If
End If
End Sub

Private Sub Form_Load()
xh = Int(Rnd * (99999999 - 1 + 1)) + 1
jc = 0
tg = 0
sc = 0
Command4.Caption = "点击后删除文本框（√）"
End Sub
