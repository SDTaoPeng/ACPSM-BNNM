Attribute VB_Name = "Module1"
'全局变量声明
Public Tfx() As Double    '投影指标   (向量,以各数据点为元素)
Public Fxy() As Double    '数据点矩阵 (第1下标为数据点序号,第2下标为数据维数)
                          '每个数据为行向量(Fxy(i,1),Fxy(i,2))
                          
Public xymax1 As Double, Sumx1 As Double, Sumy1 As Double
Public Learn_Pause As String


'NewT
'平面数据点类型定义
Public Type xy
     x As Double
     y As Double
End Type
Public CurcvsPoint(10000) As xy

Public TDataFileName As String
Public HSDataFileName As String
Public GenDataFileName As String


