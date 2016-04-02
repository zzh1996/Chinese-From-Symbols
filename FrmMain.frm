VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "字符文字V1.1"
   ClientHeight    =   5235
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   768
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox WidthChoose 
      Caption         =   "压缩"
      Height          =   255
      Left            =   10560
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复制"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox MaskCountBox 
      Height          =   300
      ItemData        =   "FrmMain.frx":0000
      Left            =   9600
      List            =   "FrmMain.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox IsBold 
      Caption         =   "加粗"
      Height          =   180
      Left            =   10560
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox FontChoose 
      Height          =   300
      ItemData        =   "FrmMain.frx":0004
      Left            =   7560
      List            =   "FrmMain.frx":0011
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1440
      Width           =   11055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "在CMD中查看"
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "文字"
      Top             =   240
      Width           =   11055
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   0
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "  by负一的平方根  2014年6月"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "模糊程度"
      Height          =   180
      Left            =   8760
      TabIndex        =   4
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AscData(127, 3) As Integer

Sub Go0()
    Dim Length As Integer
    Dim MaskCount As Integer
    Dim i As Integer, j As Integer
    Dim Result As String
    
    Pic1.Cls
    Pic1.Print Text1.Text
    
    Length = LenB(StrConv(Text1.Text, vbFromUnicode))
    'Pic1.Line (0, 0)-(Length * 8, 16), , B
    
    MaskCount = Val(MaskCountBox.Text)
    If MaskCount > 0 Then
        For i = 1 To MaskCount
            AddMask Length * 10, 16
        Next
    End If
    
    Result = ""
    
    If WidthChoose.Value = 1 Then
    
        For i = 0 To 15 Step 2
            For j = 0 To Length * 10 - 1 Step 2
                Result = Result & Chr(GetMatchAsc(Pic1.Point(j, i) Mod 256, Pic1.Point(j + 1, i) Mod 256, Pic1.Point(j, i + 1) Mod 256, Pic1.Point(j + 1, i + 1) Mod 256))
            Next
            Result = Result & vbCrLf
        Next
        
    Else
    
        For i = 0 To 15 Step 2
            For j = 0 To Length * 10 - 1
                Result = Result & Chr(GetMatchAsc(Pic1.Point(j, i) Mod 256, Pic1.Point(j, i) Mod 256, Pic1.Point(j, i + 1) Mod 256, Pic1.Point(j, i + 1) Mod 256))
            Next
            Result = Result & vbCrLf
        Next
    
    End If
    
    Text2.Text = Result
End Sub

Private Sub Command1_Click()
    Open App.Path & "\Result.txt" For Output As #1
    Print #1, Text2.Text
    Close #1
    Open App.Path & "\ViewResult.bat" For Output As #1
    Print #1, "@type " & """" & App.Path & "\Result.txt" & """"
    Print #1, "@Pause"
    Close #1
    Shell App.Path & "\ViewResult.bat", vbNormalFocus
End Sub

Sub AddMask(X As Integer, Y As Integer)
    Dim i As Integer, j As Integer
    Dim Map1() As Integer
    ReDim Map1(-1 To X + 1, -1 To Y + 1) As Integer
    For i = -1 To X + 1
        Map1(i, -1) = 255
        Map1(i, Y + 1) = 255
    Next
    For i = -1 To Y + 1
        Map1(-1, i) = 255
        Map1(X + 1, i) = 255
    Next
    For i = 0 To X
        For j = 0 To Y
            Map1(i, j) = Pic1.Point(i, j) Mod 256
        Next
    Next
    For i = 0 To X
        For j = 0 To Y
            Pic1.PSet (i, j), Int((Map1(i, j) + Map1(i, j - 1) + Map1(i, j + 1) + Map1(i - 1, j - 1) + Map1(i - 1, j) + Map1(i - 1, j + 1) + Map1(i + 1, j - 1) + Map1(i + 1, j) + Map1(i + 1, j + 1)) / 9) * RGB(1, 1, 1)
        Next
    Next
End Sub

Function GetMatchAsc(Color1 As Integer, Color2 As Integer, Color3 As Integer, Color4 As Integer) As Integer
    Dim Dist As Long, MinDist As Long, MinAsc As Integer
    MinDist = 262144
    Dim i As Integer
    For i = 32 To 126
        Dist = Abs(AscData(i, 0) - Color1) + Abs(AscData(i, 1) - Color2) + Abs(AscData(i, 2) - Color3) + Abs(AscData(i, 3) - Color4)
        If Dist < MinDist Then
            MinDist = Dist
            MinAsc = i
        End If
    Next
    GetMatchAsc = MinAsc
End Function

Private Sub Command2_Click()
    Clipboard.SetText Text2.Text
End Sub

Private Sub FontChoose_Click()
    Pic1.FontName = FontChoose.Text
    Go0
End Sub

Private Sub Form_Load()
    Pic1.FontName = "Terminal"
    Pic1.FontSize = 16
    GenAscData
    Pic1.FontSize = 12
    Text2.FontName = "Terminal"
    Text2.FontSize = 16
    FontChoose.ListIndex = 1
    Dim i As Integer
    For i = 0 To 5
        MaskCountBox.AddItem i
    Next
    MaskCountBox.ListIndex = 0
    Text1_Change
End Sub

Sub GenAscData()
    Dim i As Integer
    For i = 0 To 127
        Pic1.Cls
        Pic1.Print Chr(i)
        AscData(i, 0) = AvgColor(0, 0)
        AscData(i, 1) = AvgColor(5, 0)
        AscData(i, 2) = AvgColor(0, 10)
        AscData(i, 3) = AvgColor(5, 10)
    Next
    Pic1.Cls
End Sub

Function AvgColor(X As Integer, Y As Integer) As Integer
    Dim i As Integer, j As Integer
    Dim Sum As Long
    Sum = 0
    For i = X To X + 4
        For j = Y To Y + 9
            Sum = Sum + (Pic1.Point(i, j) Mod 256)
        Next
    Next
    AvgColor = Sum / 50
End Function

Private Sub IsBold_Click()
    Pic1.FontBold = (IsBold.Value = 1)
    Go0
End Sub

Private Sub MaskCountBox_Click()
    Go0
End Sub

Private Sub Text1_Change()
    Go0
End Sub

Private Sub WidthChoose_Click()
    Go0
End Sub
