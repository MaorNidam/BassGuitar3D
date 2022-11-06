Attribute VB_Name = "m3draw"
Option Explicit
Public Const MaxPoints = 800
Private Const parallel = 0
Private Const perspective = 1
Private dist As Single
Private projec As Byte
Private xOrg As Single
Private yOrg As Single
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private light As Vector
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Sub InitLightVector()
    light.x = 0
    light.y = 0
    light.Z = 1
End Sub
Public Function GetLightVector() As Vector
    GetLightVector = light
End Function
Public Sub LightVectorApply(ByRef M As m3Matrix)
    m3ops.m3VectorApply light, M
    m3ops.m3VectorSetLength 1, light
End Sub
Public Sub m3SetPerspective()
    projec = perspective
    If dist = 0 Then dist = 600
    
End Sub
Public Sub m3SetParallel()
    projec = parallel
End Sub
Public Function m3isparallel() As Boolean
    m3isparallel = (projec = parallel)
    
End Function
Public Function m3getdistance() As Single
    m3getdistance = dist
End Function
Public Sub m3setdistance(ByVal new_dist As Single)
    dist = new_dist
End Sub
Public Sub m3linedraw(ByRef obj As Object, ByRef P1 As m3Point, ByRef P2 As m3Point)
    Dim x1 As Long
    Dim y1 As Long
    Dim x2 As Long
    Dim y2 As Long
    Dim pt As POINTAPI
    Dim f As Single
    xOrg = -obj.ScaleLeft
    yOrg = obj.ScaleTop
    Select Case projec
        Case parallel
            x1 = xOrg + P1.x
            y1 = yOrg - P1.y
            x2 = xOrg + P2.x
            y2 = yOrg - P2.y
        Case perspective
            f = dist / (dist - P1.Z)
            x1 = xOrg + f * P1.x
            y1 = yOrg - f * P1.y
            f = dist / (dist - P2.Z)
            x2 = xOrg + f * P2.x
            y2 = yOrg - f * P2.y
    End Select
    MoveToEx obj.hdc, x1, y1, pt
    LineTo obj.hdc, x2, y2
    SetPixel obj.hdc, x2, y2, obj.ForeColor
End Sub

Public Function m3PlaneIsVisible(ByRef P1 As m3Point, ByRef P2 As m3Point, ByRef P3 As m3Point) As Boolean
    Dim Pp1 As m3Point
    Dim Pp2 As m3Point
    Dim Pp3 As m3Point
    Dim f As Double
    Dim v As Vector
    Dim u As Vector
    Dim w As Vector
    If (projec = perspective) Then
        f = dist / (dist - P1.Z)
        Pp1.x = f * P1.x
        Pp1.y = f * P1.y
        Pp1.Z = P1.Z
        
        f = dist / (dist - P2.Z)
        Pp2.x = f * P2.x
        Pp2.y = f * P2.y
        Pp2.Z = P2.Z
        
        f = dist / (dist - P3.Z)
        Pp3.x = f * P3.x
        Pp3.y = f * P3.y
        Pp3.Z = P3.Z
    Else
        Pp1.x = P1.x
        Pp1.y = P1.y
        Pp1.Z = P1.Z
        
        Pp2.x = P2.x
        Pp2.y = P2.y
        Pp2.Z = P2.Z
        
        Pp3.x = P3.x
        Pp3.y = P3.y
        Pp3.Z = P3.Z
    End If
    v = m3VectorInit(Pp1, Pp2)
    u = m3VectorInit(Pp2, Pp3)
    w = m3VectorCross(v, u)
    m3PlaneIsVisible = (w.Z >= 0.000001)
End Function
