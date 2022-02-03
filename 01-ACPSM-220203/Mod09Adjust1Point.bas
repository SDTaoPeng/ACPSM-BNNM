Attribute VB_Name = "Mod09"
'入口参数:m-顶点序号,dv-调整步长
'全局变量:V()-顶点数组 GV()

'Const Namnapp = 0.13
Public Sub Adjust1Point(ByVal m As Integer, ByVal dv As Double)
    Dim Vtemp0 As xy, Vtemp1 As xy, GVtemp As Double, n As Integer
    Dim d1 As Double, d2 As Double
    Dim i As Integer, j As Integer
    Dim DistanceofDtoVSZtemp As Double   '数据点到曲线的总距离平方（上次）
    Dim bz As Boolean
    Dim Xjiaxs As Double, Xjianxs As Double, Yjiaxs As Double, Yjianxs As Double '设4个系数
    Xjiaxs = 1#: Xjianxs = 1#: Yjiaxs = 1#: Yjianxs = 1#                         '防止溢出
    
    If V(m).X >= 0.99 Then Xjiaxs = 0
    If V(m).X <= -0.99 Then Xjianxs = 0
    
    If V(m).Y >= 0.99 Then Yjiaxs = 0
    If V(m).Y <= -0.99 Then Yjianxs = 0
    
'    If (Abs(V(m).y) >= 0.92) Then Yjiaxs = 0
'    If Abs(V(m).x <= 0.08) Then Xjianxs = 0
'    If Abs(V(m).y <= 0.08) Then Yjianxs = 0
  
  
    '保存当前顶点在Vtemp0及Vtemp1中,保存当前GV(m)在GVtemp中
    '保存DistanceofDtoVSZ在DistanceofDtoVSZtemp中
    Vtemp0.X = V(m).X: Vtemp0.Y = V(m).Y: Vtemp1.X = V(m).X: Vtemp1.Y = V(m).Y
    GVtemp = GV(m)
    DistanceofDtoVSZtemp = DistanceofDtoVSZ
    'n = m
    'If m = LBound(V) Then n = m + 1
    'If m = UBound(V) Then n = m - 1
    i = 0
    i = i + 1
    V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y - Yjianxs * dv
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)      '调整
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    i = i + 1
    V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)     '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整

    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
     i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
     i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y - Yjianxs * dv
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y + Yjiaxs * dv
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    '
    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y - Yjianxs * dv
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    d1 = MoveDirectionDistance(1): j = 1               '求出最小的
    For i = 2 To 8
        If MoveDirectionDistance(i) < d1 Then
           'd1 = MoveDirectionDistance(i): j = i
           d1 = MoveDirectionDistance(i)
           j = i
        End If
    Next i
  
  
    Dim DistanceofDtoVSZtemp1 As Double   '临时变量，上次
    Dim DistanceofDtoVSZtemp2 As Double   '临时变量，上次
    Dim DistanceofDtoVSZtemp3 As Double   '临时变量，上次

    If (DistanceofDtoVSZtemp - MoveDirectionDistance(j) > 0.002) Then
          V(m).X = MoveDirectionV(j).X: V(m).Y = MoveDirectionV(j).Y   '顶点新的位置
          Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
          Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
         
         For i = LBound(V) To UBound(V)                   
            bz = False
            If (Pi(i) - 1) >= 0.01 Then bz = True: Exit For  '角度惩罚
         Next i
         
         If (bz = True) Then
                V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y
         End If
    Else
         V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y
    End If
    Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
    Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
Adjust1Point_Exit:

End Sub



Public Sub Adjust1PointSub1(ByVal i As Integer, ByVal m As Integer, ByVal DistanceofDtoVSZtemp As Double, ByRef Vtemp1 As xy)
     Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
     Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
     MoveDirectionDistance(i) = DistanceofDtoVSZ '数据点到曲线的总距离平方
     MoveDirectionV(i).X = V(m).X: MoveDirectionV(i).Y = V(m).Y        '8个新顶点
End Sub



