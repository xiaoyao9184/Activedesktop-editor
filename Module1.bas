Attribute VB_Name = "Module1"
Option Explicit
Public CFG(9) As Byte '动态数据
Public RES() As String '动态数据


Public PictureFT As Byte, PicturePath As Byte
Public WallpaperPath$, IconPath$

Public SavePath As String '保存路径
Public SaveName As String '保存名

Public iconNO1 As Byte '现在第一个显示图标的编号
Public iconNOW As Byte '选中图标编号（相对与iconNO1）
Public KEYindex As Byte '保证选择的键值 正确写入到某个txtKey控件中
Public ICONindex As Byte '记录上个操作图标的编号(从1开始),协助完成提醒
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'获取Winsows系统文件夹

Public Sub Main()
    frmMain.Show
    '取得命令行参数
    If Len(Command()) <> 0 Then
        Dim iFileName As Variant '路径数组
        Dim iName As String '路径
        iName = Replace(Command(), Chr(34), "") '替换"为空
        iFileName = Split(iName, "\") '返回一个下标从零开始的一维数组，它包含指定数目的子字符串。
        If OpenCR(Mid(iName, 1, Len(iName) - 4) & ".cfg", Mid(iName, 1, Len(iName) - 4) & ".res", Left(iName, Len(iName) - Len(iFileName(UBound(iFileName))))) = False Then Exit Sub
        SavePath = Left(iName, Len(iName) - Len(iFileName(UBound(iFileName)))) '得到路径
        SaveName = iFileName(UBound(iFileName)) '得到名称
        frmMain.mSave.Enabled = True
        frmMain.maSave.Enabled = True
    End If
End Sub
'加载设置
Public Sub load_cfg()
Open App.Path & "\Config.cfg" For Binary As #1
Get #1, 1, PictureFT
Get #1, 2, PicturePath
Seek #1, 8
Line Input #1, WallpaperPath
Seek #1, Seek(1) + 4
Line Input #1, IconPath
Close #1
End Sub
'加载背景状态栏图片
Public Sub apply_Picture(FT)
Dim nowPath$
If FT = 1 Then
    If PicturePath = 0 Then
        nowPath = App.Path & "\Default\icon"
        frmMain.Wallpaper.Picture = LoadPicture(nowPath & "\Wallpaper.jpg")
    Else
        nowPath = IconPath
        frmMain.Wallpaper.Picture = LoadPicture(WallpaperPath)
    End If
    frmMain.Imgicon(0).Picture = LoadPicture(nowPath & "\415.gif")
    frmMain.Imgicon(1).Picture = LoadPicture(nowPath & "\404.gif")
    frmMain.Imgicon(2).Picture = LoadPicture(nowPath & "\407.gif")
    frmMain.Imgicon(3).Picture = LoadPicture(nowPath & "\473.gif")
    frmMain.Imgicon(4).Picture = LoadPicture(nowPath & "\391.gif")
    frmMain.Imgicon(5).Picture = LoadPicture(nowPath & "\457.gif")
    frmMain.Imgicon(6).Picture = LoadPicture(nowPath & "\394.gif")
    frmMain.Imgicon(8).Picture = LoadPicture(nowPath & "\416.gif")
    frmMain.Imgicon(9).Picture = LoadPicture(nowPath & "\335.gif")
End If
End Sub

'读文件
Public Function OpenCR(cfgPath As String, RESPath As String, Path As String) As Boolean
Dim i As Byte
'读CFG
    Open cfgPath For Binary As #1
    If LOF(1) <> 10 Then MsgBox "不是正确的CFG文件！无法继续。", , "警告！": OpenCR = False: Exit Function
    Seek #1, 1
    For i = 0 To 9
        Get #1, i + 1, CFG(i)
    Next
    Close #1
    frmMain.Combo1.Clear: frmMain.Combo1.Text = "请选择图标"
    For i = 1 To frmMain.Iicon.UBound
        Unload frmMain.Iicon(i)
    Next
'读RES（支持中文）
Dim TwoByte(1) As Byte, i1%, j%
ReDim Preserve RES(j)
    Open RESPath For Binary As #1
        Get #1, 1, TwoByte
        If TwoByte(0) <> 254 And TwoByte(1) <> 255 Then MsgBox "不是正确的RES文件！无法继续。", , "警告！": OpenCR = False: Exit Function
        frmMain.Combo1.AddItem "图标1"
    For i1 = 3 To LOF(1) - 1 Step 2
        Get #1, i1, TwoByte
        If TwoByte(0) = 0 And TwoByte(1) = 13 Then
            j = j + 1
        ElseIf TwoByte(0) = 0 And TwoByte(1) = 10 Then
            ReDim Preserve RES(j)
            frmMain.Combo1.AddItem "图标" & j + 1
            Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
        Else
            RES(j) = RES(j) & ChrW(CLng(TwoByte(0)) * 256 + TwoByte(1))
        End If
    Next
        Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
    Close #1
'读RES
    'Open RESPath For Input As #1
    'i = 0
    'Do Until EOF(1)
    '    ReDim Preserve RES(i)
    '    Line Input #1, RES(i)
    '    frmMain.Combo1.AddItem "图标" & i + 1
    '    Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
    '    i = i + 1
    'Loop
    'RES(0) = Mid(RES(0), 2, Len(RES(0)) - 1)
    'Close #1
Dim a As Byte
'显示数据
    frmMain.txtX = CFG(0)
    frmMain.txtY = CFG(1)
    frmMain.txtZ = CFG(2)
    frmMain.optType(CFG(3)).Value = True
    frmMain.optPath(CFG(4)).Value = True
    frmMain.optTf(CFG(5)).Value = True
    frmMain.txtMuch = CFG(6)
    frmMain.txtKey(0).Text = CFG(7)
    frmMain.txtKey(1).Text = CFG(8)
    frmMain.txtKey(2).Text = CFG(9)
'加载图片
    If LoadP(Path & "ActdeskIcon\", 255) = False Then MsgBox "没有 " & Path & "ActdeskIcon\ 目录！", , "提示！": OpenCR = False: Exit Function
    'Iicon(0)为背景
    If LoadP(Path & "ActdeskIcon\DOCK.GIF", 0) = False Then OpenCR = False: Exit Function
    frmMain.Iicon(0).Left = CFG(0) * 15
    frmMain.Iicon(0).Top = CFG(1) * 15
    'Iicon(*)*为图标编号（从1开始）
    For a = 1 To frmMain.Iicon.UBound
        If LoadP(Path & "\ActdeskIcon\ICON" & a & ".GIF", a) = False Then OpenCR = False: Exit Function
    Next
    'Iicon(UBound)为光标
    Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
    If LoadP(Path & "\ActdeskIcon\CURSOR.GIF", frmMain.Iicon.UBound) = False Then OpenCR = False: Exit Function
'定位预览位置
    Call ShowICON(1, 1)
    iconNO1 = 1
    iconNOW = 1
    OpenCR = True
End Function

Public Function LoadP(Path As String, NO As Byte) As Boolean
    If NO = 255 Then LoadP = IIf(Dir(Path) = "", False, True): Exit Function
    If Dir(Path) = "" Then
        If MsgBox("没有 " & Path & "图片！是否继续加载？", vbYesNo, "提示！") = vbYes Then
            frmMain.Iicon(NO).Picture = LoadPicture(App.Path & "\Default\ActdeskIcon\NO.GIF")
        Else
            LoadP = False: Exit Function
        End If
    Else
        frmMain.Iicon(NO).Picture = LoadPicture(Path)
    End If
    LoadP = True
End Function
'写文件
Public Sub SaveCR(Path As String)
    Dim i As Byte
'写CFG
    Open Path & ".cfg" For Binary As #2
    For i = 0 To 9
        Put #2, i + 1, CFG(i)
    Next
    Close #2
'RES组织成一个String
    Dim all As String
    For i = 0 To UBound(RES)
        all = all & RES(i) & IIf(i = UBound(RES), Null, Chr(13) & Chr(10))
    Next
'RES组织成一个Byte
    Dim ALLbyte() As Byte
    ReDim ALLbyte(Len(all) * 2 - 1)
    For i = 0 To Len(all) - 1
        ALLbyte(i * 2) = AscW(Mid(all, i + 1, 1)) \ 256
        ALLbyte(i * 2 + 1) = AscW(Mid(all, i + 1, 1)) Mod 256
    Next
'写RES
    Open Path & ".res" For Binary As #2
        Put #2, 1, 254
        Put #2, 2, 255
        Put #2, 3, ALLbyte 'all
    Close #2
End Sub



'调整图标显示
Public Sub ShowICON(No1 As Byte, MNo As Byte) '（图标编号(显示的图标 中的第一个)，光标(相对与 显示的图标 的编号)）
Dim i As Byte
'关闭显示图标
For i = 1 To frmMain.Iicon.UBound
    frmMain.Iicon(i).Visible = False
Next
'显示图标
frmMain.Iicon(No1).Left = CFG(0) * 15
frmMain.Iicon(No1).Top = (CFG(1) + 15) * 15
frmMain.Iicon(No1).Visible = True
frmMain.Iicon(No1).ZOrder 0
For i = No1 + 1 To No1 + CFG(6) - 1
    If CFG(3) = 0 Then
        frmMain.Iicon(i).Top = frmMain.Iicon(i - 1).Top
        frmMain.Iicon(i).Left = frmMain.Iicon(i - 1).Left + frmMain.Iicon(i - 1).Width + CFG(2) * 15
    Else
        frmMain.Iicon(i).Left = frmMain.Iicon(i - 1).Left
        frmMain.Iicon(i).Top = frmMain.Iicon(i - 1).Top + frmMain.Iicon(i - 1).Height + CFG(2) * 15
    End If
    frmMain.Iicon(i).Visible = True
    frmMain.Iicon(i).ZOrder 0
Next
'光标图片
frmMain.Iicon(frmMain.Iicon.UBound).Left = frmMain.Iicon(No1 + MNo - 1).Left
frmMain.Iicon(frmMain.Iicon.UBound).Top = frmMain.Iicon(No1 + MNo - 1).Top
frmMain.Iicon(frmMain.Iicon.UBound).Visible = True
frmMain.Iicon(frmMain.Iicon.UBound).ZOrder 0
'选择图标编辑
frmMain.Combo1.ListIndex = No1 + MNo - 1 - 1 '（No1 + MNo  -1）代表图标编号（从1开始）
End Sub



