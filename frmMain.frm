VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Activedesktop 编辑器"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   885
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdB 
      Caption         =   "→"
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   32
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdB 
      Caption         =   "↓"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   30
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton CmdB 
      Caption         =   "↑"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   29
      Top             =   3840
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   7320
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   2
      Left            =   6600
      MaxLength       =   3
      TabIndex        =   27
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   1
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   26
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   0
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   25
      Top             =   2160
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "单个图标设置"
      Height          =   1095
      Left            =   3360
      TabIndex        =   18
      Top             =   3120
      Width           =   3975
      Begin VB.TextBox txtPath 
         Height          =   270
         Left            =   1440
         TabIndex        =   22
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Opt 
         Caption         =   "启动ELF"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Opt 
         Caption         =   "发送代码"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "代码/路径"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3360
      TabIndex        =   17
      Text            =   "请选择图标"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtMuch 
      Height          =   270
      Left            =   6240
      MaxLength       =   3
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "显示设置"
      Height          =   855
      Left            =   6240
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
      Begin VB.OptionButton optTf 
         Caption         =   "开"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optTf 
         Caption         =   "关"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "面板盘符"
      Height          =   855
      Left            =   4800
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
      Begin VB.OptionButton optPath 
         Caption         =   "A/"
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPath 
         Caption         =   "C/"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "显示类型"
      Height          =   855
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
      Begin VB.OptionButton optType 
         Caption         =   "纵向"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optType 
         Caption         =   "横向"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtZ 
      Height          =   270
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtY 
      Height          =   270
      Left            =   4200
      MaxLength       =   3
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtX 
      Height          =   270
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Wallpaper 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   360
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3300
      ScaleWidth      =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2640
      Begin VB.Image Iicon 
         Height          =   495
         Index           =   0
         Left            =   240
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":370F
         Top             =   0
         Width           =   330
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   1
         Left            =   330
         Picture         =   "frmMain.frx":37A5
         Top             =   0
         Width           =   210
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   2
         Left            =   540
         Picture         =   "frmMain.frx":3829
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   3
         Left            =   780
         Picture         =   "frmMain.frx":38A7
         Top             =   0
         Width           =   195
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   4
         Left            =   975
         Picture         =   "frmMain.frx":3914
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   5
         Left            =   1245
         Picture         =   "frmMain.frx":399F
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   6
         Left            =   1515
         Picture         =   "frmMain.frx":3A4F
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   7
         Left            =   1755
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   8
         Left            =   2025
         Picture         =   "frmMain.frx":3AFA
         Top             =   0
         Width           =   285
      End
      Begin VB.Image Imgicon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   9
         Left            =   2310
         Picture         =   "frmMain.frx":3B8A
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.CommandButton CmdB 
      Caption         =   "←"
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   31
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "退出组合键："
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "+          +"
      Height          =   135
      Left            =   5400
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "图标显示个数"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "图标间距"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "起始坐标：X Y"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Menu mField 
      Caption         =   "文件(&F)"
      Begin VB.Menu mNew 
         Caption         =   "新建(&N)"
      End
      Begin VB.Menu mOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu maSave 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mf1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mSet 
         Caption         =   "设置(&S)"
      End
      Begin VB.Menu mAbout 
         Caption         =   "关于本程序(&A)..."
      End
   End
   Begin VB.Menu mnuMNU 
      Caption         =   " "
      Visible         =   0   'False
      Begin VB.Menu mEdit 
         Caption         =   "编辑"
         Visible         =   0   'False
      End
      Begin VB.Menu mOpenPath 
         Caption         =   "打开目录"
      End
      Begin VB.Menu mOpenRes 
         Caption         =   "打开RES"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim L As String, i As Integer

'读键值列表
Open App.Path & "\Default\KEY.cfg" For Input As #1
    Do Until EOF(1)
        Line Input #1, L
        List1.AddItem L
    Loop
Close #1
'初值
iconNO1 = 1: iconNOW = 1
'调整
List1.Top = txtKey(0).Top + txtKey(0).Height
List1.Left = txtKey(0).Left
List1.Visible = False

load_cfg
apply_Picture (PictureFT)
End Sub










'菜单
Private Sub mNew_Click()
    If OpenCR(App.Path & "\Default\Actdesk.cfg", App.Path & "\Default\Actdesk.res", App.Path & "\Default\") = False Then Exit Sub
    SavePath = ""
End Sub
Private Sub mOpen_Click()
    CommonDialog1.Filter = "Activedesktop 配置文件|*.res;*.cfg"
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    If OpenCR(Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - 4) & ".cfg", Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - 4) & ".res", Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))) = False Then Exit Sub
    SavePath = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
    SaveName = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
End Sub
Private Sub mSave_Click()
    If SavePath = "" Then Call maSave_Click: Exit Sub
    SaveCR (SavePath & SaveName)
End Sub
Private Sub maSave_Click()
    CommonDialog1.Filter = "Activedesktop 配置文件|*.res;*.cfg"
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    If LoadP(Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle)) & "ActdeskIcon\", 255) = False Then
        Dim fs, f
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(IIf(SavePath = "", App.Path & "\Default\ActdeskIcon\", SavePath & "ActdeskIcon\"))
        f.Copy Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
    End If
    SavePath = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
    SaveName = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
    SaveCR (SavePath & SaveName)
End Sub
Private Sub mExit_Click()
    If MsgBox("保存修改？", vbOKCancel, "提醒") = 1 Then Call mSave_Click
    End
End Sub
Private Sub mSet_Click()
    frmSetting.Show
End Sub
Private Sub mAbout_Click()
    frmAbout.Show
End Sub





'限制输入
Private Sub txtX_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtX.Text, 1, txtX.SelStart) & Chr(KeyAscii) & Mid(txtX.Text, txtX.SelStart + 1)) > 176 Then KeyAscii = 0
End Sub
Private Sub txtY_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtY.Text, 1, txtY.SelStart) & Chr(KeyAscii) & Mid(txtY.Text, txtY.SelStart + 1)) > 220 Then KeyAscii = 0
End Sub
Private Sub txtZ_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtZ.Text, 1, txtZ.SelStart) & Chr(KeyAscii) & Mid(txtZ.Text, txtZ.SelStart + 1)) > IIf(CFG(3) = 0, 176, 220) Then KeyAscii = 0
End Sub
Private Sub txtMuch_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtMuch.Text, 1, txtMuch.SelStart) & Chr(KeyAscii) & Mid(txtMuch.Text, txtMuch.SelStart + 1)) > 255 Then KeyAscii = 0
If CInt(Mid(txtMuch.Text, 1, txtMuch.SelStart) & Chr(KeyAscii) & Mid(txtMuch.Text, txtMuch.SelStart + 1)) > Iicon.UBound - 1 Then KeyAscii = 0
End Sub
Private Sub txtKey_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtKey(Index).Text, 1, txtKey(Index).SelStart) & Chr(KeyAscii) & Mid(txtKey(Index).Text, txtKey(Index).SelStart + 1)) > 255 Then KeyAscii = 0
End Sub
Private Sub txtPath_KeyPress(KeyAscii As Integer)
If Opt(0).Value = True Then
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
Else
    If KeyAscii = 92 Or KeyAscii = 58 Or KeyAscii = 63 Or KeyAscii = 42 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 124 Or KeyAscii = 34 Then KeyAscii = 0
End If
End Sub


'及时改变
Private Sub txtX_Change()
If txtX.Text <> "" Then CFG(0) = CInt(txtX.Text): frmMain.Iicon(0).Left = CFG(0) * 15
End Sub
Private Sub txtY_Change()
If txtY.Text <> "" Then CFG(1) = CInt(txtY.Text): frmMain.Iicon(0).Top = CFG(1) * 15
End Sub
Private Sub txtZ_Change()
If txtZ.Text <> "" Then CFG(2) = CInt(txtZ.Text): ShowICON iconNO1, iconNOW
End Sub
Private Sub txtMuch_Change()
If txtMuch.Text <> "" Then CFG(6) = CInt(txtMuch.Text): iconNOW = 1: ShowICON iconNO1, iconNOW
End Sub
Private Sub optType_Click(Index As Integer)
CFG(3) = Index
ShowICON iconNO1, iconNOW
End Sub
Private Sub optPath_Click(Index As Integer)
CFG(4) = Index
End Sub
Private Sub optTf_Click(Index As Integer)
CFG(5) = Index
End Sub
Private Sub txtKey_Change(Index As Integer)
If txtKey(Index).Text <> "" Then CFG(Index + 7) = CInt(txtKey(Index).Text)
End Sub
Private Sub Opt_Click(Index As Integer)
If Index = IIf(Left(RES(Combo1.ListIndex), 3) = "ELF", 1, 0) Then Exit Sub '如果选择的相同，则退出
txtPath.Text = IIf(Index, "/b/ELF/", 1000)
End Sub
Private Sub txtPath_Change()
If Right(txtPath.Text, 1) = " " Then txtPath.Text = Left(txtPath.Text, Len(txtPath.Text) - 1) '消除空格后缀
txtPath.ToolTipText = txtPath.Text
RES(Combo1.ListIndex) = IIf(Opt(1).Value, "ELF=", "icon" & Combo1.ListIndex + 1 & "=") & txtPath.Text '更新数据
End Sub


'组合键选择
Private Sub txtKey_Click(Index As Integer)
KEYindex = Index
List1.Visible = False
If txtKey(Index).Text = "" Then Exit Sub
If CInt(txtKey(Index).Text) < 69 Then: List1.ListIndex = txtKey(Index).Text: KEYindex = Index: List1.Visible = True
End Sub
'组合键列表
Private Sub List1_Click()
Dim i As Byte
'检测相同的键值
For i = 0 To 2
    If KEYindex <> i Then
        If txtKey(i).Text = List1.ListIndex Then
            MsgBox "第 " & KEYindex + 1 & " 个键与第 " & i + 1 & " 个键设置了相同的键值！", , "不能设置相同的键值！": Exit Sub
        End If
    End If
Next
txtKey(KEYindex).Text = List1.ListIndex
List1.Visible = False
KEYindex = 3
End Sub
Private Sub Form_Click()
If List1.Visible = True Then txtKey(KEYindex) = List1.ListIndex: List1.Visible = False: KEYindex = 3
End Sub
'选择图标
Private Sub Combo1_Click()
'提醒：TXT为空
    If txtPath.Text = "" And Combo1.ListIndex + 1 <> ICONindex And ICONindex <> 0 Then
        MsgBox "没有填写图标 " & ICONindex & " 的 " & IIf(Opt(1).Value, "路径", "事件代码")
        Combo1.ListIndex = ICONindex - 1
        Exit Sub
    End If
    ICONindex = Combo1.ListIndex + 1
'显示指定 编号图标(从0开始)内容
    Opt(IIf(Left(RES(Combo1.ListIndex), 3) = "ELF", 1, 0)) = True
    txtPath.Text = Mid(RES(Combo1.ListIndex), IIf(Left(RES(Combo1.ListIndex), 3) = "ELF", 5, 7))

'及时预览
If Iicon(0).Visible = False Then Exit Sub '隐藏不运行
If Combo1.ListIndex + 1 <= CFG(6) + iconNO1 - 1 And Combo1.ListIndex + 1 >= iconNO1 Then
    iconNOW = Combo1.ListIndex + 1 - iconNO1 + 1
ElseIf Combo1.ListIndex + 1 < iconNO1 Then
    iconNO1 = Combo1.ListIndex + 1
    iconNOW = 1
ElseIf Combo1.ListIndex + 1 > CFG(6) + iconNO1 - 1 Then
    iconNO1 = Combo1.ListIndex + 1 - CFG(6) + 1
    iconNOW = CFG(6)
End If
Call ShowICON(iconNO1, iconNOW)
End Sub
'滚动预览
Private Sub CmdB_Click(Index As Integer)
Dim i%
If Index = IIf(CFG(3), 1, 3) Then '向下滚动
    If Iicon(0).Visible = False Then Exit Sub '隐藏不运行
    If iconNOW + 1 > CFG(6) Then
        If iconNO1 = Iicon.UBound - 1 - CFG(6) + 1 Then '是最后一个图标,转向第1个
            iconNO1 = 1
            iconNOW = 1
        Else
            iconNO1 = iconNO1 + 1
        End If
    Else
        iconNOW = iconNOW + 1
    End If
    Call ShowICON(iconNO1, iconNOW)
ElseIf Index = IIf(CFG(3), 0, 2) Then '向上滚动
    If Iicon(0).Visible = False Then Exit Sub
    If iconNOW - 1 < 1 Then
        If iconNO1 = 1 Then
            iconNO1 = Iicon.UBound - 1 - CFG(6) + 1
            iconNOW = CFG(6)
        Else
            iconNO1 = iconNO1 - 1
        End If
    Else
        iconNOW = iconNOW - 1
    End If
    Call ShowICON(iconNO1, iconNOW)
ElseIf Index = IIf(CFG(3), 3, 0) Then '显示
    Iicon(0).Visible = True
    Iicon(Iicon.UBound).Visible = True
    For i = iconNO1 To iconNO1 + CFG(6) - 1
        Iicon(i).Visible = True
    Next
ElseIf Index = IIf(CFG(3), 2, 1) Then '隐藏
    For i = 0 To Iicon.UBound
        Iicon(i).Visible = False
    Next
End If
End Sub


Private Sub Iicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then frmMain.PopupMenu mnuMNU, vbPopupMenuLeftAlign, X + Iicon(Index).Left + Wallpaper.Left, Y + Iicon(Index).Top + Wallpaper.Top
End Sub

Private Sub mEdit_Click()
Dim s$
ChDrive Left(IIf(SavePath = "", App.Path & "\Default\ActdeskIcon\Icon\", SavePath & "\ActdeskIcon\Icon\"), 1)
s = IIf(SavePath = "", App.Path & "\Default\ActdeskIcon\Icon\", SavePath & "\ActdeskIcon\Icon\")
ChDir s
s = Combo1.ListIndex + 1 & ".gif"
On Error Resume Next
Shell "C:\WINDOWS\system32\mspaint.exe " & s, 4
If Err.Number = 53 Then MsgBox "Win图片编辑器不存在！", , "提示！"
End Sub
Private Sub mOpenPath_Click()
Dim sTmp As String * 200, Length As Long
Length = GetWindowsDirectory(sTmp, 200)
Shell Left(sTmp, Length) & "\explorer.exe " & IIf(SavePath = "", App.Path & "\Default\ActdeskIcon\", SavePath & "\ActdeskIcon\"), 4
End Sub
Private Sub mOpenres_Click()
Dim s$
s = IIf(SavePath = "", App.Path & "\Default\Actdesk.res", SavePath & SaveName & ".res")
On Error Resume Next
Shell "C:\WINDOWS\system32\notepad.exe " & s, 4
If Err.Number = 53 Then MsgBox "WinTXT编辑器不存在！", , "提示！"
End Sub
