VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   Caption         =   "设置"
   ClientHeight    =   4425
   ClientLeft      =   4185
   ClientTop       =   3330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   5730
   Begin VB.CommandButton cmdapply 
      Caption         =   "应用"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "加载外部图片（包括背景和状态栏图标）"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame fraPicture 
      Caption         =   "图片目录"
      Height          =   2775
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtWallpaper 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "状态栏目录"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optPictureSuff 
         Caption         =   "默认"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPictureOther 
         Caption         =   "其他目录"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "背景图片"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "选择目录下任何一个GIF文件即可指定此目录为状态栏目录"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "程序目录\Default\icon\"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "必须同时指定背景图片和状态栏图标目录"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Close #3
load_cfg

chkPicture.Value = PictureFT
If PictureFT = 0 Then
    optPictureSuff.Enabled = False
    optPictureOther.Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
    cmdapply.Enabled = False
End If
If PicturePath = 0 Then
    optPictureSuff.Value = True
    txtWallpaper.Text = App.Path & "\icon\Wallpaper.jpg"
    txtIcon.Text = App.Path & "\icon"
Else
    optPictureOther.Value = True
    txtWallpaper.Text = WallpaperPath
    txtIcon.Text = IconPath
End If
End Sub
Private Sub chkPicture_Click()
PictureFT = chkPicture.Value
If chkPicture.Value = 1 Then
    optPictureSuff.Enabled = True
    optPictureOther.Enabled = True
    If optPictureSuff.Value = False Then
        cmdWallpaper.Enabled = True
        cmdIcon.Enabled = True
    End If
    cmdapply.Enabled = True
Else
    optPictureSuff.Enabled = False
    optPictureOther.Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
    cmdapply.Enabled = False
End If
End Sub
Private Sub optPictureSuff_Click()
If optPictureSuff.Value = True Then
    PicturePath = 0
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
End If
End Sub
Private Sub optPictureOther_Click()
If optPictureOther.Value = True Then
    PicturePath = 1
    cmdWallpaper.Enabled = True
    cmdIcon.Enabled = True
End If
End Sub
Private Sub cmdWallpaper_Click()
CommonDialog1.Filter = "BMP(*.bmp)|*.bmp|GIF(*.gif)|*.gif|JPEG(*.jpg)|*.jpg"
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtWallpaper.Text = CommonDialog1.FileName
    txtWallpaper.ToolTipText = CommonDialog1.FileName
    WallpaperPath = CommonDialog1.FileName
End Sub
Private Sub cmdIcon_Click()
CommonDialog1.Filter = "任何一个状态栏图标(*.gif)|*.gif"
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtIcon.Text = CurDir()
    txtIcon.ToolTipText = CurDir()
    IconPath = CurDir()
End Sub
Private Sub cmdSave_Click()
Open App.Path & "\Config.cfg" For Binary As #1
Put #1, 1, CByte(PictureFT)
Put #1, 2, CByte(PicturePath)
Put #1, 4, WallpaperPath & Chr(13) & Chr(10)
Put #1, , IconPath & Chr(13) & Chr(10)
Close #1
End Sub
Private Sub cmdapply_Click()
apply_Picture (PictureFT)
End Sub
