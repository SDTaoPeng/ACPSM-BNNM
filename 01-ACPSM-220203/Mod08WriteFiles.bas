Attribute VB_Name = "Mod08"
'-----------------------------------------------------------------------------------------------
'
Public Sub WriteFile(t() As Double, data() As xy, FileName As String)
    '声明本子程序使用的临时变量
    Dim Iw As Integer
    Dim Sw As String
    '打开文件供写入(若原有数据,会被清空)
    Open FileName For Output As #1      '打开文件
    '逐行写入
    For Iw = LBound(t) To UBound(t)
        Sw = Format$(t(Iw), "0.0000000000") & " "
        
        If data(Iw).X >= 0 Then
           Sw = Sw & "+" & Format$(data(Iw).X, "0.0000000000") & " "
        Else
           Sw = Sw & Format$(data(Iw).X, "#0.0000000000") & " "
        End If
        
        If data(Iw).Y >= 0 Then
           Sw = Sw & "+" & Format$(data(Iw).Y, "0.0000000000")
        Else
           Sw = Sw & Format$(data(Iw).Y, "#0.0000000000")
        End If
        
        Print #1, Sw             '写入1行(Sw)
    Next Iw
    '写入xymax,Sumx,Sumy,
        Sw = Format$(xymax, "0.0000000000") & " "
        
        If Sumx >= 0 Then
           Sw = Sw & "+" & Format$(Sumx, "0.0000000000") & " "
        Else
           Sw = Sw & Format$(Sumx, "#0.0000000000") & " "
        End If
        
        If Sumy >= 0 Then
           Sw = Sw & "+" & Format$(Sumy, "0.0000000000") & " "
        Else
           Sw = Sw & Format$(Sumy, "#0.0000000000") & " "
        End If
    Print #1, Sw              '写入1行(Sw)
    '关闭文件
    Close #1
End Sub
