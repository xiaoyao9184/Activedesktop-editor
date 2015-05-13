Attribute VB_Name = "Module1"
Option Explicit
Public CFG(9) As Byte '��̬����
Public RES() As String '��̬����


Public PictureFT As Byte, PicturePath As Byte
Public WallpaperPath$, IconPath$

Public SavePath As String '����·��
Public SaveName As String '������

Public iconNO1 As Byte '���ڵ�һ����ʾͼ��ı��
Public iconNOW As Byte 'ѡ��ͼ���ţ������iconNO1��
Public KEYindex As Byte '��֤ѡ��ļ�ֵ ��ȷд�뵽ĳ��txtKey�ؼ���
Public ICONindex As Byte '��¼�ϸ�����ͼ��ı��(��1��ʼ),Э���������
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'��ȡWinsowsϵͳ�ļ���

Public Sub Main()
    frmMain.Show
    'ȡ�������в���
    If Len(Command()) <> 0 Then
        Dim iFileName As Variant '·������
        Dim iName As String '·��
        iName = Replace(Command(), Chr(34), "") '�滻"Ϊ��
        iFileName = Split(iName, "\") '����һ���±���㿪ʼ��һά���飬������ָ����Ŀ�����ַ�����
        If OpenCR(Mid(iName, 1, Len(iName) - 4) & ".cfg", Mid(iName, 1, Len(iName) - 4) & ".res", Left(iName, Len(iName) - Len(iFileName(UBound(iFileName))))) = False Then Exit Sub
        SavePath = Left(iName, Len(iName) - Len(iFileName(UBound(iFileName)))) '�õ�·��
        SaveName = iFileName(UBound(iFileName)) '�õ�����
        frmMain.mSave.Enabled = True
        frmMain.maSave.Enabled = True
    End If
End Sub
'��������
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
'���ر���״̬��ͼƬ
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

'���ļ�
Public Function OpenCR(cfgPath As String, RESPath As String, Path As String) As Boolean
Dim i As Byte
'��CFG
    Open cfgPath For Binary As #1
    If LOF(1) <> 10 Then MsgBox "������ȷ��CFG�ļ����޷�������", , "���棡": OpenCR = False: Exit Function
    Seek #1, 1
    For i = 0 To 9
        Get #1, i + 1, CFG(i)
    Next
    Close #1
    frmMain.Combo1.Clear: frmMain.Combo1.Text = "��ѡ��ͼ��"
    For i = 1 To frmMain.Iicon.UBound
        Unload frmMain.Iicon(i)
    Next
'��RES��֧�����ģ�
Dim TwoByte(1) As Byte, i1%, j%
ReDim Preserve RES(j)
    Open RESPath For Binary As #1
        Get #1, 1, TwoByte
        If TwoByte(0) <> 254 And TwoByte(1) <> 255 Then MsgBox "������ȷ��RES�ļ����޷�������", , "���棡": OpenCR = False: Exit Function
        frmMain.Combo1.AddItem "ͼ��1"
    For i1 = 3 To LOF(1) - 1 Step 2
        Get #1, i1, TwoByte
        If TwoByte(0) = 0 And TwoByte(1) = 13 Then
            j = j + 1
        ElseIf TwoByte(0) = 0 And TwoByte(1) = 10 Then
            ReDim Preserve RES(j)
            frmMain.Combo1.AddItem "ͼ��" & j + 1
            Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
        Else
            RES(j) = RES(j) & ChrW(CLng(TwoByte(0)) * 256 + TwoByte(1))
        End If
    Next
        Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
    Close #1
'��RES
    'Open RESPath For Input As #1
    'i = 0
    'Do Until EOF(1)
    '    ReDim Preserve RES(i)
    '    Line Input #1, RES(i)
    '    frmMain.Combo1.AddItem "ͼ��" & i + 1
    '    Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
    '    i = i + 1
    'Loop
    'RES(0) = Mid(RES(0), 2, Len(RES(0)) - 1)
    'Close #1
Dim a As Byte
'��ʾ����
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
'����ͼƬ
    If LoadP(Path & "ActdeskIcon\", 255) = False Then MsgBox "û�� " & Path & "ActdeskIcon\ Ŀ¼��", , "��ʾ��": OpenCR = False: Exit Function
    'Iicon(0)Ϊ����
    If LoadP(Path & "ActdeskIcon\DOCK.GIF", 0) = False Then OpenCR = False: Exit Function
    frmMain.Iicon(0).Left = CFG(0) * 15
    frmMain.Iicon(0).Top = CFG(1) * 15
    'Iicon(*)*Ϊͼ���ţ���1��ʼ��
    For a = 1 To frmMain.Iicon.UBound
        If LoadP(Path & "\ActdeskIcon\ICON" & a & ".GIF", a) = False Then OpenCR = False: Exit Function
    Next
    'Iicon(UBound)Ϊ���
    Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
    If LoadP(Path & "\ActdeskIcon\CURSOR.GIF", frmMain.Iicon.UBound) = False Then OpenCR = False: Exit Function
'��λԤ��λ��
    Call ShowICON(1, 1)
    iconNO1 = 1
    iconNOW = 1
    OpenCR = True
End Function

Public Function LoadP(Path As String, NO As Byte) As Boolean
    If NO = 255 Then LoadP = IIf(Dir(Path) = "", False, True): Exit Function
    If Dir(Path) = "" Then
        If MsgBox("û�� " & Path & "ͼƬ���Ƿ�������أ�", vbYesNo, "��ʾ��") = vbYes Then
            frmMain.Iicon(NO).Picture = LoadPicture(App.Path & "\Default\ActdeskIcon\NO.GIF")
        Else
            LoadP = False: Exit Function
        End If
    Else
        frmMain.Iicon(NO).Picture = LoadPicture(Path)
    End If
    LoadP = True
End Function
'д�ļ�
Public Sub SaveCR(Path As String)
    Dim i As Byte
'дCFG
    Open Path & ".cfg" For Binary As #2
    For i = 0 To 9
        Put #2, i + 1, CFG(i)
    Next
    Close #2
'RES��֯��һ��String
    Dim all As String
    For i = 0 To UBound(RES)
        all = all & RES(i) & IIf(i = UBound(RES), Null, Chr(13) & Chr(10))
    Next
'RES��֯��һ��Byte
    Dim ALLbyte() As Byte
    ReDim ALLbyte(Len(all) * 2 - 1)
    For i = 0 To Len(all) - 1
        ALLbyte(i * 2) = AscW(Mid(all, i + 1, 1)) \ 256
        ALLbyte(i * 2 + 1) = AscW(Mid(all, i + 1, 1)) Mod 256
    Next
'дRES
    Open Path & ".res" For Binary As #2
        Put #2, 1, 254
        Put #2, 2, 255
        Put #2, 3, ALLbyte 'all
    Close #2
End Sub



'����ͼ����ʾ
Public Sub ShowICON(No1 As Byte, MNo As Byte) '��ͼ����(��ʾ��ͼ�� �еĵ�һ��)�����(����� ��ʾ��ͼ�� �ı��)��
Dim i As Byte
'�ر���ʾͼ��
For i = 1 To frmMain.Iicon.UBound
    frmMain.Iicon(i).Visible = False
Next
'��ʾͼ��
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
'���ͼƬ
frmMain.Iicon(frmMain.Iicon.UBound).Left = frmMain.Iicon(No1 + MNo - 1).Left
frmMain.Iicon(frmMain.Iicon.UBound).Top = frmMain.Iicon(No1 + MNo - 1).Top
frmMain.Iicon(frmMain.Iicon.UBound).Visible = True
frmMain.Iicon(frmMain.Iicon.UBound).ZOrder 0
'ѡ��ͼ��༭
frmMain.Combo1.ListIndex = No1 + MNo - 1 - 1 '��No1 + MNo  -1������ͼ���ţ���1��ʼ��
End Sub



