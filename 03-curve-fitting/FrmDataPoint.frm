VERSION 5.00
Begin VB.Form FrmDataPoint 
   Caption         =   "基于主曲线方法的特性曲线数值拟合(取数据点)         "
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16950
   Icon            =   "FrmDataPoint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "FrmDataPoint.frx":0442
   Moveable        =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   16950
   Begin VB.PictureBox PicData 
      Height          =   9645
      Left            =   3480
      ScaleHeight     =   9585
      ScaleMode       =   0  'User
      ScaleWidth      =   11240.36
      TabIndex        =   42
      Top             =   240
      Width           =   11055
   End
   Begin VB.CommandButton ImportRawdata 
      Caption         =   "导入Rawdata"
      Height          =   495
      Left            =   15120
      TabIndex        =   41
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton SaveRawdata 
      Caption         =   "存Rawdata"
      Height          =   495
      Left            =   15120
      TabIndex        =   40
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox TxtQN 
      Height          =   285
      Index           =   1
      Left            =   1890
      TabIndex        =   24
      Text            =   "100"
      Top             =   4440
      Width           =   795
   End
   Begin VB.TextBox TxtState 
      Height          =   255
      Left            =   1140
      TabIndex        =   21
      Text            =   "TxtState"
      Top             =   10320
      Width           =   13515
   End
   Begin VB.Frame Frame1 
      Caption         =   "取数据点操作区域"
      Height          =   10065
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3045
      Begin VB.CommandButton CmdQXJZ 
         Caption         =   "倾斜校正"
         Height          =   495
         Left            =   2220
         TabIndex        =   33
         Top             =   1740
         Width           =   555
      End
      Begin VB.Frame Frame4 
         Caption         =   "取数据点"
         Height          =   4785
         Left            =   180
         TabIndex        =   25
         Top             =   5160
         Width           =   2685
         Begin VB.Frame Frame5 
            Caption         =   "输入线条特征(例XL=90)"
            Height          =   885
            Left            =   120
            TabIndex        =   36
            Top             =   1110
            Width           =   2295
            Begin VB.TextBox TxtPara 
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Text            =   "AA=0"
               Top             =   210
               Width           =   2025
            End
            Begin VB.CommandButton CmdSave 
               Caption         =   "存数据点"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   540
               Width           =   1965
            End
         End
         Begin VB.FileListBox FileCH 
            Height          =   1170
            Left            =   120
            TabIndex        =   32
            Top             =   2580
            Width           =   2355
         End
         Begin VB.TextBox TxtFileName 
            Height          =   315
            Left            =   30
            TabIndex        =   31
            Text            =   "TxtFileName"
            Top             =   3960
            Width           =   2445
         End
         Begin VB.CommandButton CmdCLS 
            Caption         =   "清屏"
            Height          =   375
            Left            =   1260
            TabIndex        =   30
            Top             =   4320
            Width           =   915
         End
         Begin VB.CommandButton CmdDrawDataPoint 
            Caption         =   "画数据点"
            Height          =   375
            Left            =   210
            TabIndex        =   29
            Top             =   4320
            Width           =   915
         End
         Begin VB.CommandButton CmdZBscale 
            Caption         =   "坐标尺度"
            Height          =   255
            Left            =   210
            TabIndex        =   28
            Top             =   210
            Width           =   2055
         End
         Begin VB.CommandButton CmdYQ 
            Caption         =   "预取2-去密集点"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   27
            Top             =   780
            Width           =   2085
         End
         Begin VB.CommandButton CmdYQ 
            Caption         =   "预取1-去除黑色点"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   480
            Width           =   2085
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "说明:鼠标左键变红色,右键变黑色.将通过预取1去除."
            Height          =   435
            Left            =   150
            TabIndex        =   39
            Top             =   2070
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 坐标变换数据采集"
         Height          =   2865
         Left            =   180
         TabIndex        =   7
         Top             =   2220
         Width           =   2775
         Begin VB.TextBox TxtQN 
            Height          =   285
            Index           =   0
            Left            =   1650
            TabIndex        =   23
            Text            =   "1000"
            Top             =   1710
            Width           =   795
         End
         Begin VB.CommandButton CmdParaSave 
            Caption         =   "保存初始参数"
            Height          =   285
            Left            =   300
            TabIndex        =   20
            Top             =   2490
            Width           =   2000
         End
         Begin VB.TextBox TxtXY0 
            Height          =   255
            Index           =   5
            Left            =   810
            TabIndex        =   19
            Text            =   "屏幕值"
            Top             =   2130
            Width           =   765
         End
         Begin VB.TextBox TxtXY0 
            Height          =   255
            Index           =   4
            Left            =   1980
            TabIndex        =   17
            Text            =   "40"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TxtXY0 
            Height          =   225
            Index           =   1
            Left            =   1980
            TabIndex        =   16
            Text            =   "Y0"
            Top             =   1020
            Width           =   555
         End
         Begin VB.TextBox TxtYPM 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Text            =   "TxtYPM"
            Top             =   570
            Width           =   855
         End
         Begin VB.TextBox TxtXPM 
            Height          =   285
            Left            =   360
            TabIndex        =   14
            Text            =   "TxtXPM"
            Top             =   600
            Width           =   765
         End
         Begin VB.TextBox TxtXY0 
            Height          =   255
            Index           =   3
            Left            =   870
            TabIndex        =   13
            Text            =   "屏幕值"
            Top             =   1740
            Width           =   705
         End
         Begin VB.TextBox TxtXY0 
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   11
            Text            =   "200"
            Top             =   1350
            Width           =   675
         End
         Begin VB.TextBox TxtXY0 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   9
            Text            =   "X0"
            Top             =   1020
            Width           =   555
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "在右图中单击鼠标获得屏幕坐标"
            Height          =   255
            Left            =   90
            TabIndex        =   34
            Top             =   300
            Width           =   2625
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   2610
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Label LblQZ 
            BackColor       =   &H00FF80FF&
            Caption         =   "Yn11B"
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   18
            Top             =   2130
            Width           =   615
         End
         Begin VB.Label LblQZ 
            BackColor       =   &H00FF80FF&
            Caption         =   "XQ11B"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   1710
            Width           =   615
         End
         Begin VB.Label LblQZ 
            Caption         =   "输入Q11 n11"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   10
            Top             =   1350
            Width           =   1095
         End
         Begin VB.Label LblQZ 
            BackColor       =   &H00FF80FF&
            Caption         =   "单击取X0,Y0"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   1050
            Width           =   1005
         End
      End
      Begin VB.TextBox TxtEZH 
         Height          =   315
         Left            =   570
         TabIndex        =   6
         Text            =   "二值化阈值"
         Top             =   1800
         Width           =   825
      End
      Begin VB.PictureBox Slider1 
         Height          =   1905
         Left            =   120
         MousePointer    =   4  'Icon
         ScaleHeight     =   1845
         ScaleWidth      =   225
         TabIndex        =   5
         Top             =   270
         Width           =   285
      End
      Begin VB.CommandButton CmdEZH 
         Caption         =   "二值化"
         Height          =   465
         Left            =   1470
         TabIndex        =   4
         Top             =   1740
         Width           =   705
      End
      Begin VB.TextBox TxtXY 
         Height          =   285
         Left            =   1110
         TabIndex        =   3
         Text            =   "TxtXY"
         Top             =   1410
         Width           =   1785
      End
      Begin VB.TextBox TxtImageOK 
         Height          =   285
         Left            =   510
         TabIndex        =   2
         Text            =   "TxtImageOK"
         Top             =   1020
         Width           =   2355
      End
      Begin VB.FileListBox FilePic 
         Height          =   630
         Left            =   480
         TabIndex        =   1
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "点颜色"
         Height          =   255
         Left            =   510
         TabIndex        =   35
         Top             =   1440
         Width           =   585
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "过程信息"
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   10320
      Width           =   855
   End
End
Attribute VB_Name = "FrmDataPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'数据声明
Option Explicit                                '检查未经声明的变量
Dim YQ As String
Dim XPM, YPM As Single   ' 坐标
Dim X0, Y0 As Single     '
Dim XQ11A, XQ11B As Single    '
Dim Yn11A, Yn11B As Single    '
Dim PQ11, Pn11 As Single    '
Dim ZBscale As Byte

' Dim XS() As Boolean '定义存放图象像素的数组
' Dim XSZB() As Boolean '备用
 
 Dim XS() As Long '定义存放图象像素的数组
 Dim XSZB() As Long '备用

Private Sub CmdQXJZ_Click()
   Dim P1, P2 As Integer ' 坐标限
   Dim X, Y As Integer '图象点坐标
    Dim xp, yp As Integer '移正图象后的点坐标
    Dim xpd, ypd As Double '移正图象后的点坐标()
  
   
   PicData.ScaleMode = vbPixels '图象按像素分辨
    '取图象坐标限
    P1 = PicData.ScaleWidth
    P2 = PicData.ScaleHeight
    '按图象大小分配点空间(2值)
    ReDim XS(0 To P1 + 5, 0 To P2 + 5)
    ReDim XSZB(0 To P1 + 5, 0 To P2 + 5)
   
   
'获取图象信息---------------------------------------------------------------------
    '扫描图象,获取并保存所有点信息
    For X = 0 To P1
        For Y = 0 To P2
           XS(X, Y) = PicData.Point(X, Y)
           '若当前点是黑色的,点阵数组元素值为true
'           If PicData.Point(X, Y) = &H0 Then
'              XS(X, Y) = True
'           '否则为false
'           Else
'              XS(X, Y) = False
'           End If
        
        Next Y
    Next X
    
    '检测是否记录下所有点信息(图案变蓝)
    For X = 0 To P1
        For Y = 0 To P2
            PicData.PSet (X, Y), vbWhite
              '黑点变蓝
'            If XS(X, Y) Then
'                PicData.PSet (X, Y), vbWhite  'vbYellow '    'vbBlack
'            End If
          Next Y
    Next X
    

'移正图象--------------------------------------------------------------------------
    '备份移动前的图象信息
    For X = 0 To P1
        For Y = 0 To P2
              XSZB(X, Y) = XS(X, Y)
        Next Y
    Next X
    'xs中存放移动后的图象信息
   
    
    For X = 1 To P1 - 2
        For Y = 1 To P2

              
              
              
              xpd = X + (616 - Y) * Sqr(0.0001 / (1 + 0.0001))
              
              ypd = Y + ((616 - Y) - (616 - Y) * Sqr(1 / (1 + 0.0001)))
              
              xp = Int(xpd + 0.49 * ((X Mod 6) / 6))
              yp = Int(ypd + 0.49 * ((X Mod 6) / 6))
              
              If xp <= 0 Then xp = 0
              If yp <= 0 Then yp = 0

              XS(X, Y) = XSZB(xp, yp)
              
        Next Y
    Next X
    '重画
    For X = 1 To P1 - 5
        For Y = 1 To P2
            PicData.PSet (X, Y), XS(X, Y)
'           If XS(X, Y) Then
'              PicData.PSet (X, Y), vbBlack 'vbBlue 'vbGreen  'vbYellow    'vbBlack
'           End If
        Next Y
    Next X
 
End Sub

'窗体初始化
Public Sub Form_Initialize()
    FilePic.Path = "..\SLJTU"       '设置图片数据文件路径
    'Slider1.Value = 86              '设置滑动条二值化初值
    TxtEZH.Text = 86
    YQ = "不取"
    CmdEZH.Enabled = False          '
    ZBscale = 1 '坐标尺度初值
    
    FileCH.Path = "..\SRCdata"       '  "..\BPCSdata"      '原始数据
    FileCH.FileName = "*.ABC"
    '图象按像素分辨
    PicData.ScaleMode = vbPixels
    TxtState.Text = "取点操作顺序：选取型号->二值化->[坐标变换]->鼠标移动选线->存数据点->复核": TxtState.Refresh ' 过程信息
End Sub

Private Sub CmdYQ_Click(Index As Integer) '"预取"
    Dim Px, Py As Single    '坐标限
    Dim X, Y As Single      '坐标
    Dim s As String
    Px = PicData.ScaleWidth: Py = PicData.ScaleHeight  '取图象坐标限
   Select Case Index
      Case 0   '"预取1"
        TxtState.Text = "预取1:正在去除黑色点，请稍候.......": TxtState.Refresh ' 过程信息
        For X = 0 To Px
           For Y = 0 To Py
           '若当前点
            If PicData.Point(X, Y) <> vbRed Then PicData.PSet (X, Y), vbWhite
           Next Y
        Next X
        TxtState.Text = "预取1:正在去除黑色点，请稍候.......已经完成，可以进行其他操作": TxtState.Refresh ' 过程信息
      Case 1   '"预取2"
        TxtState.Text = "预取2:正在去除密集点，请稍候.......": TxtState.Refresh ' 过程信息
        For X = 0 To Px
           For Y = 0 To Py
           '若当前点
            If PicData.Point(X, Y) = vbRed Then
               'PicData.PSet (X, Y - 1), vbWhite
               PicData.PSet (X, Y + 1), vbWhite
               PicData.PSet (X - 1, Y - 1), vbWhite
               'PicData.PSet (X - 1, Y), vbWhite
               'PicData.PSet (X - 1, Y + 1), vbWhite
               PicData.PSet (X + 1, Y - 1), vbWhite
               'PicData.PSet (X + 1, Y), vbWhite
               'PicData.PSet (X + 1, Y + 1), vbWhite
             End If
           Next Y
        Next X
        TxtState.Text = "预取2:正在去除密集点，请稍候.......已经完成，可以进行其他操作": TxtState.Refresh ' 过程信息
   End Select

End Sub

Private Sub CmdZBscale_Click()
    ZBscale = ZBscale + 1
    If ZBscale > 8 Then ZBscale = 1
    CmdZBscale.Caption = "坐标尺度=" & ZBscale
End Sub

'重画文件选择
Private Sub FileCH_Click()
   TxtFileName.Text = FileCH.FileName
End Sub



'在图片文件选择框中单击选择图形文件
Private Sub FilePic_DblClick()  '双击选择图形文件(双击-防止误动作)
    Dim FileName, Sw As String, s1 As String, S2 As String
    
    Dim i As Byte
  
    
  TxtImageOK.Text = "图:" & FilePic.FileName: TxtImageOK.Refresh        '显示选择的图形文件名
  '在图片框中显示图片
  PicData.Picture = LoadPicture(FilePic.Path & "\" & FilePic.FileName)
  '图象按像素分辨
  PicData.ScaleMode = vbPixels
  '把图片缩放的图片框中
  PicData.PaintPicture PicData.Picture, 0, 0, PicData.ScaleWidth, PicData.ScaleHeight
  CmdEZH.Enabled = True
  '
  i = InStr(FilePic.FileName, ".")
  FileName = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & "??=*.txt"
  
  FileCH.Path = "..\SRCdata"       '  "..\BPCSdata"      '原始数据
  
  FileCH.FileName = FileName   'HL180A194
  
  TxtState.Text = "FilePic_Click() 选择图形文件 "  ' 过程信息
End Sub






'滑动条决定二值化阈值
Private Sub Slider1_Click()
   'TxtEZH.Text = Slider1.Value
   TxtEZH.Text = 86
End Sub

'二值化
Private Sub CmdEZH_Click()
    Dim Px, Py As Single    '坐标限
    Dim X, Y As Single      '坐标
    Dim s As String         '
    TxtState.Text = "特别提示：二值化过程较慢，请勿移动鼠标或其他操作，否则会引起失误.....": TxtState.Refresh        '过程信息
    '取图象坐标限
    Px = PicData.ScaleWidth
    Py = PicData.ScaleHeight
    s = Hex$(TxtEZH.Text) & Hex$(TxtEZH.Text) & Hex$(TxtEZH.Text)
    
     For X = 0 To Px
        For Y = 0 To Py
           
           '若当前点是黑色的,点阵数组元素值为true
           If PicData.Point(X, Y) < ("&H" & s) Then
             
             PicData.PSet (X, Y), vbBlack
           '否则为false
           Else
           
              PicData.PSet (X, Y), vbWhite
              
           End If
        
        Next Y
    Next X
    TxtState.Text = "特别提示：二值化过程较慢，请勿移动鼠标或其他操作，否则会引起失误......已经完成，可以进行其他操作！": TxtState.Refresh        '过程信息
End Sub

'取X0,Y0,XQ11A,XQ11B,Yn11A,Yn11B
Private Sub LblQZ_Click(Index As Integer)
   Select Case Index
      Case 0
        TxtXY0(0).Text = TxtXPM.Text: X0 = TxtXY0(0).Text
        TxtXY0(1).Text = TxtYPM.Text: Y0 = TxtXY0(1).Text
      Case 1
        'TxtXY0(2).Text = TxtXPM.Text: XQ11A = TxtXY0(2).Text
      Case 2
        TxtXY0(3).Text = TxtXPM.Text: XQ11B = TxtXY0(3).Text
      Case 3
        'TxtXY0(4).Text = TxtYPM.Text: Yn11A = TxtXY0(4).Text
      Case 4
        TxtXY0(5).Text = TxtYPM.Text: Yn11B = TxtXY0(5).Text
  End Select
End Sub

'在图片框中单击鼠标事件
Private Sub PicData_CLICK()
     TxtXPM.Text = XPM: TxtYPM.Text = YPM
End Sub


'保存初始参数
Private Sub CmdParaSave_Click()
    Dim FileName, Sw As String, s1 As String, S2 As String
    Dim i As Byte
    Dim i1, j1, i2, j2 As Single
    If (Val(TxtXPM.Text) = 0 Or Val(TxtXY0(5).Text) = 0) Then MsgBox "没有数据需要存储": Exit Sub
    i = InStr(FilePic.FileName, ".")
    FileName = "..\SRCdata\" & Left$(FilePic.FileName, i) & "txt"
    TxtState.Text = "存盘文件名：" & FileName        '过程信息
    '打开文件供写入(若原有数据,会被清空)
    Open FileName For Output As #1      '打开文件
    '逐行写入
    s1 = Format$(X0, "0000.000000"): S2 = Format$(Y0, "0000.000000")         '1.坐标原点
    Sw = " 原点X0  =" & s1 & " Y0  =" & S2
    Print #1, Sw             '写入1行
    
    i1 = Val(TxtXY0(2).Text): j1 = Val(TxtXY0(4).Text)                       '2.Q11 n11原点
    s1 = Format$(i1, "0000.000000"): S2 = Format$(j1, "0000.000000")
    Sw = " 原点Q11 =" & s1 & " n11 =" & S2
    Print #1, Sw             '写入1行
    
    i2 = TxtQN(0).Text                                                       '3.PQ11 PN11
    PQ11 = (XQ11B - X0) / (i2 - i1)
    j2 = TxtQN(1).Text
    Pn11 = (Yn11B - Y0) / (j2 - j1)
    s1 = Format$(PQ11, "0000.000000"): S2 = Format$(Pn11, "0000.000000")
    Sw = " 比例PQ11=" & s1 & " Pn11=" & S2
    
    Print #1, Sw             '写入1行
     '关闭文件
    Close #1
End Sub

Private Sub SaveRawdata_Click()   '存数据点，然后预处理后，再导入
    Dim FileName, Sw As String
    Dim Px, Py As Single    '坐标
    Dim i As Byte
    Dim X, Y As Single     '坐标
    i = InStr(FilePic.FileName, ".")
    FileName = "..\Rawdata\" & Left$(FilePic.FileName, i) & "txt"
    TxtState.Text = "存盘文件名：" & FileName        '过程信息
    Open FileName For Output As #1      '打开文件
    '逐行写入
    Px = PicData.ScaleWidth: Py = PicData.ScaleHeight
    For X = 0 To Px
       For Y = 0 To Py
           '若当前点
           If PicData.Point(X, Y) = vbRed Then
               Sw = Format$(X, "0000.000000") & " " & Format$(Y, "0000.000000")
              Print #1, Sw             '写入1行
              TxtState.Text = "特别提示：正在存储原始数据，请勿移动鼠标或其他操作，否则会引起失误": TxtState.Refresh
           End If
         Next Y
    Next X
    Close #1
    TxtState.Text = "存储完成!": TxtState.Refresh     '过程信息
End Sub

Private Sub ImportRawdata_Click()   '导入数据点，将预处理后的数据点导入
    Dim FileName As String
    Dim fileline() As String
    Dim i As Byte
    Dim N1, Nd As Integer
    Dim Px, Py As Single    '坐标限
    Dim X1, Y1 As Single     '坐标

    '先取参数---------------------------------------------------------------------------
    i = InStr(FilePic.FileName, ".")
    FileName = "..\Rawdata\" & Left$(FilePic.FileName, i) & "txt"
    TxtState.Text = "读文件：" & FileName & "中参数": TxtState.Refresh      '过程信息
    Open FileName For Input As #2          '打开文件
    Nd = 0                           '文件总行数初值=0
    Do Until EOF(2)
        Nd = Nd + 1
        ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
        Line Input #2, fileline(Nd)      '读一行—>最新行
    Loop
    Close #2
    For N1 = 1 To Nd
        X1 = Left$(fileline(N1), 12)
        Y1 = Right$(fileline(N1), 12)
        PicData.PSet (X1, Y1), vbRed
    Next N1
    TxtState.Text = "导入数据完成!已显示": TxtState.Refresh     '过程信息
End Sub

Private Sub CmdSave_Click()   '存数据点
    Dim FileName, FileName1, Sw, Sw1 As String, s1 As String, S2 As String
    Dim fileline() As String
    Dim i, j As Byte
    Dim N1, Nd As Integer
    Dim i1, j1, i2, j2 As Single
    Dim Px, Py As Single    '坐标限
    Dim X, Y, X1, Y1 As Single     '坐标
    Dim Number As Long
    Dim Number11 As Long
    Dim ConvertedArray(1 To 2, 1 To 100) As Single  '存预处理后的点集
    
    'If (LTrim(RTrim(TxtPara.Text)) = "单线含义") Then GoTo CmdSave_Click_Error1
    
    '----------------------------------先取参数-------------------------------------
    i = InStr(FilePic.FileName, ".")
    FileName = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & ".txt"
    TxtState.Text = "读文件：" & FileName & "中参数": TxtState.Refresh      '过程信息
    Open FileName For Input As #2          '打开文件
    Nd = 0                           '文件总行数初值=0
    Do Until EOF(2)
        Nd = Nd + 1
        ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
        Line Input #2, fileline(Nd)      '读一行—>最新行
    Loop
    Close #2
    i = InStr(fileline(1), "="): X0 = Mid$(fileline(1), i + 1, 11): Y0 = Right$(fileline(1), 11)
    TxtState.Text = "(X0,Y0)=(" & X0 & "," & Y0 & ")": TxtState.Refresh    '过程信息
    i = InStr(fileline(3), "="): PQ11 = Mid$(fileline(3), i + 1, 11): Pn11 = Right$(fileline(3), 12)
    TxtState.Text = "(PQ11,PN11)=(" & PQ11 & "," & Pn11 & ")": TxtState.Refresh     '过程信息
    '------------------------------------------------------------------------------
       
    '-----------------------------老版本--直接将pixturebox中的红色点写入-------------
    i = InStr(FilePic.FileName, ".")
    Sw = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & LTrim(RTrim(TxtPara.Text))
    FileName = Sw & ".txt"
    TxtState.Text = "存盘文件名：" & FileName: TxtState.Refresh       '过程信息
    '打开文件供写入(若原有数据,会被清空)
    Open FileName For Output As #1      '打开文件
    '逐行写入数据
        
    '取图象坐标限
    Px = PicData.ScaleWidth: Py = PicData.ScaleHeight
    For X = 0 To Px
       For Y = 0 To Py
    '       '若当前点
           If PicData.Point(X, Y) = vbRed Then
               X1 = (X - X0) / (PQ11 * 1000): Y1 = (Y - Y0) / (Pn11 * 100)
               Sw = Format$(X1, "0000.000000") & " " & Format$(Y1, "0000.000000")
              Print #1, Sw             '写入1行
           End If
         Next Y
    Next X
    '关闭文件
    Close #1
    TxtState.Text = "存储完成!": TxtState.Refresh     '过程信息
    '------------------------------------------------------------------------------

    
    '-----------------------------新版本--导入处理后的点集--------------------------
    'j = InStr(FilePic.FileName, ".")
    'Sw = "..\Covertdata\FB150A101-coverted.txt"
    'FileName = Sw
    'TxtState.Text = "存盘文件名：" & FileName: TxtState.Refresh       '过程信息
    '打开文件供写入(若原有数据,会被清空)
   ' Open FileName For Input As #2          '打开文件
    'Nd = 0                           '文件总行数初值=0
    'Do Until EOF(2)
    '    Nd = Nd + 1
    '    ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
    '    Line Input #2, fileline(Nd)      '读一行—>最新行
   ' Loop
   ' Close #2
    
  '  Number = Nd
  '  For N1 = 1 To Nd
  '      X1 = Left$(fileline(N1), 12)
  '      Y1 = Right$(fileline(N1), 12)
  '      ConvertedArray(1, N1) = X1
  '      ConvertedArray(2, N1) = Y1
  '      PicData.PSet (X1, Y1), vbRed
  '  Next N1
    
    '---------------------------------打开文件--------------------------------------
   ' j1 = InStr(FilePic.FileName, ".")
   ' Sw1 = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & LTrim(RTrim(TxtPara.Text))
   ' FileName1 = Sw1 & ".txt"
  '  Open FileName1 For Output As #1      '打开文件
    '逐行写入数据
    '取图象坐标限
  '  For Number11 = 1 To Nd
           '若当前点
 '       X = ConvertedArray(1, Number11)
  '      Y = ConvertedArray(2, Number11)
  '      X1 = (X - X0) / (PQ11 * 1000): Y1 = (Y - Y0) / (Pn11 * 100)
  '      Sw = Format$(X1, "0000.000000") & " " & Format$(Y1, "0000.000000")
  '      Print #1, Sw             '写入1行
 ''   Next Number11
    
    
    '关闭文件
   ' Close #1
  '  TxtState.Text = "存储完成!": TxtState.Refresh     '过程信息
    '------------------------------------------------------------------------------
    
    '------------------------将存储的红色点回画，附成黑色点进行显示-------------------
    'TxtState.Text = "开始回读重画!": TxtState.Refresh     '过程信息
    'i = InStr(FilePic.FileName, ".")
    'Sw = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & LTrim(RTrim(TxtPara.Text))
    'FileName = Sw & ".txt"
    'Open FileName For Input As #2          '打开文件
    '   Nd = 0                           '文件总行数初值=0
    '  Do Until EOF(2)
    '     Nd = Nd + 1
    '     ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
    '     Line Input #2, fileline(Nd)      '读一行—>最新行
    '  Loop
    '  Close #2
    '  For N1 = 1 To Nd
    '      X1 = Left$(fileline(N1), 12): X = (PQ11 * 1000) * X1 + X0
    '      Y1 = Right$(fileline(N1), 12): Y = (Pn11 * 100) * Y1 + Y0
    '       PicData.PSet (X, Y), vbBlack
    ''  Next N1
    '   '更新文件列表
    '   i = InStr(FilePic.FileName, ".")
    '   FileName = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & "??=*.txt"
    '   FileCH.Path = "..\SRCdata"       '  "..\BPCSdata"      '原始数据
    '   FileCH.FileName = FileName
    '   FileCH.Refresh
    '   TxtState.Text = "重画完成，单线存储列表已经更新!": TxtState.Refresh     '过程信息
    ' GoTo CmdSave_Click_Exit
    '------------------------------------------------------------------------
'CmdSave_Click_Error1:
'     TxtState.Text = "无合适的数据存储!"
CmdSave_Click_Exit:

End Sub


Private Sub CmdCLS_Click()    '清屏

    Dim Px, Py As Single    '坐标限
    Dim X, Y As Single      '坐标
        '
    '取图象坐标限
    Px = PicData.ScaleWidth
    Py = PicData.ScaleHeight
    
    For X = 0 To Px
       For Y = 0 To Py
           
              PicData.PSet (X, Y), vbWhite
          
          Next Y
    Next X
End Sub

Private Sub CmdDrawDataPoint_Click()   '画数据点
    Dim FileName, Sw As String, s1 As String, S2 As String
    Dim fileline() As String
    Dim i As Byte
    Dim N1, Nd As Integer
    Dim i1, j1, i2, j2 As Single
    Dim Px, Py As Single    '坐标限
    Dim X, Y, X1, Y1 As Single     '坐标
        If Len(RTrim(TxtImageOK.Text)) < 12 Then GoTo CmdDrawDataPoint_Click_Error1
        '先取参数---------------------------------------------------------------------------
        On Error GoTo CmdDrawDataPoint_Click_Error1
        i = InStr(FilePic.FileName, ".")
        FileName = "..\SRCdata\" & Left$(FilePic.FileName, i - 1) & ".txt"
        TxtState.Text = "读文件：" & FileName & "中参数": TxtState.Refresh      '过程信息
        Open FileName For Input As #2          '打开文件
        Nd = 0                           '文件总行数初值=0
        Do Until EOF(2)
           Nd = Nd + 1
           ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
           Line Input #2, fileline(Nd)      '读一行—>最新行
        Loop
        Close #2
        i = InStr(fileline(1), "="): X0 = Mid$(fileline(1), i + 1, 11): Y0 = Right$(fileline(1), 11)
        TxtState.Text = "(X0,Y0)=(" & X0 & "," & Y0 & ")": TxtState.Refresh    '过程信息
        i = InStr(fileline(3), "="): PQ11 = Mid$(fileline(3), i + 1, 11): Pn11 = Right$(fileline(3), 12)
        TxtState.Text = "(PQ11,PN11)=(" & PQ11 & "," & Pn11 & ")": TxtState.Refresh     '过程信息
       '先取参数---------------------------------------------------------------------------
    
    
    'Sw = Right$(TxtFileName, (Len(TxtFileName) - 2))
    On Error GoTo CmdDrawDataPoint_Click_Error1
    Sw = RTrim(LTrim(TxtFileName))
    FileName = "..\SRCdata\" & Sw
    TxtState.Text = "开始取" & FileName & "数据点,进行重画!": TxtState.Refresh     '过程信息
    Open FileName For Input As #2          '打开文件
       Nd = 0                           '文件总行数初值=0
      Do Until EOF(2)
         Nd = Nd + 1
         ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
         Line Input #2, fileline(Nd)      '读一行—>最新行
      Loop
      Close #2
      For N1 = 1 To Nd
          X1 = Left$(fileline(N1), 12): X = (PQ11 * 1000) * X1 + X0
          Y1 = Right$(fileline(N1), 12): Y = (Pn11 * 100) * Y1 + Y0
          TxtState.Text = X & Y: TxtState.Refresh
          PicData.PSet (X, Y), vbRed
      Next N1
     TxtState.Text = "重画" & FileName & "完成!": TxtState.Refresh     '过程信息
     GoTo CmdDrawDataPoint_Click_Exit
CmdDrawDataPoint_Click_Error1:
    TxtState.Text = "线条选择:" & "没有" & FileName & ",不能重画!"
    TxtState.Refresh
CmdDrawDataPoint_Click_Exit:
End Sub




'在图片框中移动鼠标事件
Private Sub PicData_MouseMove(button As Integer, shift As Integer, X As Single, Y As Single)
    XPM = X: YPM = Y                                                      '保存当前坐标到(XPM,YPM)
    TxtXY.Text = "(" & Format$(X, "000") & "," & Format$(Y, "000") & ")"  '显示当前坐标
    TxtXY.Text = TxtXY.Text & "=" & Hex$(PicData.Point(X, Y))             '显示当前点颜色值
    If (button = 1) Then                                    '按住鼠标右键
       Call PiontXYC(PicData, X, Y, vbBlack, vbRed, ZBscale)
       TxtState.Text = "您正在按住鼠标左键进行取点": TxtState.Refresh
    TxtState.Refresh
    End If
    If (button = 2) Then                                 '按住鼠标左键
       Call PiontXYC(PicData, X, Y, vbRed, vbBlack, ZBscale)
       TxtState.Text = "您正在按住鼠标左键进行擦点": TxtState.Refresh
    End If
    If (button = 4) Then                                 '按住鼠标中键
       PicData.PSet (X, Y), vbRed
       TxtState.Text = "您正在按住鼠标中键进行补点": TxtState.Refresh
    End If
End Sub

'功能:将图片框中(X,Y)点Drta范围内颜色为Color1的点变为颜色为Color2的
'参数:Pic--图片框；（X, Y）--坐标点；Color1，Color2--Color1-->Color2,Drta--范围
'说明：
Private Sub PiontXYC(ByVal Pic As PictureBox, X As Single, Y As Single, _
                     ByVal Color1 As Long, ByVal Color2 As Long, Drta As Byte)
   Dim xTemp As Single, yTemp As Single
   For xTemp = X - Drta To X + Drta
       For yTemp = Y - Drta To Y + Drta
       If Pic.Point(xTemp, yTemp) = Color1 Then Pic.PSet (xTemp, yTemp), Color2
     Next yTemp
   Next xTemp
End Sub












