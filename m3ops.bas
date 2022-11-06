Attribute VB_Name = "m3ops"
Option Explicit
Public Const Pie = 3.14159265358979
Public Type m3Point
    x As Double
    y As Double
    Z As Double
End Type
Public Type m3Matrix
    mat(3, 2) As Double
End Type
Public Type Vector
    x As Double
    y As Double
    Z As Double
    End Type
    
Public Function m3VectorInit(ByRef P1 As m3Point, ByRef P2 As m3Point) As Vector
    Dim v As Vector
    v.x = P2.x - P1.x
    v.y = P2.y - P1.y
    v.Z = P2.Z - P1.Z
    m3VectorInit = v
End Function

Public Function m3VectorLength(ByRef v As Vector) As Double
       m3VectorLength = Sqr(v.x * v.x + v.y * v.y + v.Z * v.Z)
End Function

Public Sub m3VectorSetLength(ByVal l As Double, ByRef v As Vector)
    l = l / m3VectorLength(v)
    v.x = v.x * l
    v.y = v.y * l
    v.Z = v.Z * l
End Sub

Public Function m3VectorSum(ByRef v As Vector, ByRef u As Vector) As Vector
    Dim w As Vector
    w.x = v.x + u.x
    w.y = v.y + u.y
    w.Z = v.Z + u.Z
    m3VectorSum = w
End Function

Public Sub m3VectorApply(ByRef v As Vector, ByRef m As m3Matrix)
    Dim u As Vector
    u.x = m.mat(0, 0) * v.x + m.mat(1, 0) * v.y + m.mat(2, 0) * v.Z
    u.y = m.mat(0, 1) * v.x + m.mat(1, 1) * v.y + m.mat(2, 1) * v.Z
    u.Z = m.mat(0, 2) * v.x + m.mat(1, 2) * v.y + m.mat(2, 2) * v.Z
    v = u
End Sub

Public Function m3VectorToString(ByRef v As Vector) As String
    Dim s As String
    s = Format(v.x, "0.00") & vbTab & Format(v.y, "0.00") & vbTab & Format(v.Z, "0.00") & 0
    m3VectorToString = s
End Function

Public Function m3PointAddVector(ByRef v As Vector, ByRef p As m3Point) As m3Point
    Dim P1 As m3Point
    P1.x = p.x + v.x
    P1.y = p.y + v.y
    P1.Z = p.Z + v.Z
    m3PointAddVector = P1
End Function

Public Function m3PointInit(ByVal x As Double, ByVal y As Double, ByVal Z As Double) As m3Point
    Dim p As m3Point
    p.x = x
    p.y = y
    p.Z = Z
    m3PointInit = p
End Function

Public Function m3PointToString(ByRef m As m3Point) As String
    m3PointToString = Format(m.x, "0.00") & vbTab & _
                      Format(m.y, "0.00") & vbTab & _
                      Format(m.Z, "0.00") & vbTab & 1
End Function

Public Function m3MatrixToString(ByRef m As m3Matrix) As String
    Dim s As String
    Dim i As Double
    Dim j As Double
    For i = 0 To 3
        For j = 0 To 2
            s = s & Format(m.mat(i, j), "0.00") & vbTab
        Next j
        s = s & i \ 3 & vbCrLf
    Next i
    m3MatrixToString = s
End Function


Public Function m3MatrixRndInit(ByVal lobound As Double, ByVal upbound As Double) As m3Matrix
    Dim i As Double
    Dim j As Double
    Dim RandomNumber As Double
    Dim m As m3Matrix
    Randomize
    For i = 0 To 3
        For j = 0 To 2
            RandomNumber = lobound + (upbound - lobound + 1) * Rnd
            m.mat(i, j) = RandomNumber
        Next j
    Next i
    m3MatrixRndInit = m
End Function

Public Function m3Identity() As m3Matrix
    Dim i As Double
    Dim j As Double
    Dim m As m3Matrix
    For i = 0 To 3
        For j = 0 To 2
            m.mat(i, j) = 0
        Next j
        If i < 3 Then
            m.mat(i, i) = 1
        End If
    Next i
    m3Identity = m
End Function

Public Function m3MatMultiply(ByRef a As m3Matrix, ByRef B As m3Matrix) As m3Matrix
    Dim c As m3Matrix
    c.mat(0, 0) = a.mat(0, 0) * B.mat(0, 0) + a.mat(0, 1) * B.mat(1, 0) + a.mat(0, 2) * B.mat(2, 0)
    c.mat(0, 1) = a.mat(0, 0) * B.mat(0, 1) + a.mat(0, 1) * B.mat(1, 1) + a.mat(0, 2) * B.mat(2, 1)
    c.mat(0, 2) = a.mat(0, 0) * B.mat(0, 2) + a.mat(0, 1) * B.mat(1, 2) + a.mat(0, 2) * B.mat(2, 2)

    c.mat(1, 0) = a.mat(1, 0) * B.mat(0, 0) + a.mat(1, 1) * B.mat(1, 0) + a.mat(1, 2) * B.mat(2, 0)
    c.mat(1, 1) = a.mat(1, 0) * B.mat(0, 1) + a.mat(1, 1) * B.mat(1, 1) + a.mat(1, 2) * B.mat(2, 1)
    c.mat(1, 2) = a.mat(1, 0) * B.mat(0, 2) + a.mat(1, 1) * B.mat(1, 2) + a.mat(1, 2) * B.mat(2, 2)
    
    c.mat(2, 0) = a.mat(2, 0) * B.mat(0, 0) + a.mat(2, 1) * B.mat(1, 0) + a.mat(2, 2) * B.mat(2, 0)
    c.mat(2, 1) = a.mat(2, 0) * B.mat(0, 1) + a.mat(2, 1) * B.mat(1, 1) + a.mat(2, 2) * B.mat(2, 1)
    c.mat(2, 2) = a.mat(2, 0) * B.mat(0, 2) + a.mat(2, 1) * B.mat(1, 2) + a.mat(2, 2) * B.mat(2, 2)
    
    c.mat(3, 0) = a.mat(3, 0) * B.mat(0, 0) + a.mat(3, 1) * B.mat(1, 0) + a.mat(3, 2) * B.mat(2, 0) + B.mat(3, 0)
    c.mat(3, 1) = a.mat(3, 0) * B.mat(0, 1) + a.mat(3, 1) * B.mat(1, 1) + a.mat(3, 2) * B.mat(2, 1) + B.mat(3, 1)
    c.mat(3, 2) = a.mat(3, 0) * B.mat(0, 2) + a.mat(3, 1) * B.mat(1, 2) + a.mat(3, 2) * B.mat(2, 2) + B.mat(3, 2)
    
    m3MatMultiply = c
End Function

Public Sub m3Pointapply(ByRef p As m3Point, a As m3Matrix)
    Dim x As Double
    Dim y As Double
    Dim Z As Double
    x = a.mat(0, 0) * p.x + a.mat(1, 0) * p.y + a.mat(2, 0) * p.Z + a.mat(3, 0)
    y = a.mat(0, 1) * p.x + a.mat(1, 1) * p.y + a.mat(2, 1) * p.Z + a.mat(3, 1)
    Z = a.mat(0, 2) * p.x + a.mat(1, 2) * p.y + a.mat(2, 2) * p.Z + a.mat(3, 2)
    p.x = x
    p.y = y
    p.Z = Z
End Sub

Public Function m3Scale(ByVal sx As Double, ByVal sy As Double, ByVal sz As Double) As m3Matrix
    Dim m As m3Matrix
    m = m3Identity
    m.mat(0, 0) = sx
    m.mat(1, 1) = sy
    m.mat(2, 2) = sz
    m3Scale = m
End Function
Public Function m3Translate(ByVal Tx As Double, ByVal Ty As Double, ByVal Tz As Double) As m3Matrix
    Dim m As m3Matrix
    m = m3Identity
    m.mat(3, 0) = Tx
    m.mat(3, 1) = Ty
    m.mat(3, 2) = Tz
    m3Translate = m
End Function

Public Function m3RotateY(ByVal zav As Double) As m3Matrix
    Dim m As m3Matrix
    m = m3Identity
    m.mat(0, 0) = Cos(zav)
    m.mat(0, 2) = -Sin(zav)
    m.mat(2, 0) = Sin(zav)
    m.mat(2, 2) = Cos(zav)
    m3RotateY = m
End Function

Public Function m3RotateZ(ByVal zav As Double) As m3Matrix
    Dim m As m3Matrix
    m = m3Identity
    m.mat(0, 0) = Cos(zav)
    m.mat(0, 1) = -Sin(zav)
    m.mat(1, 0) = Sin(zav)
    m.mat(1, 1) = Cos(zav)
    m3RotateZ = m
End Function

Public Function m3RotateX(ByVal zav As Double) As m3Matrix
    Dim m As m3Matrix
    m = m3Identity
    m.mat(1, 1) = Cos(zav)
    m.mat(1, 2) = Sin(zav)
    m.mat(2, 1) = -Sin(zav)
    m.mat(2, 2) = Cos(zav)
    m3RotateX = m
End Function

    Public Function m3PsCenter(ByVal nP As Double, ByRef p() As m3Point) As m3Point
    Dim i As Double
    Dim cP As m3Point
    For i = 0 To nP - 1
        cP.x = cP.x + p(i).x
        cP.y = cP.y + p(i).y
        cP.Z = cP.Z + p(i).Z
    Next i
    cP.x = cP.x / nP
    cP.y = cP.y / nP
    cP.Z = cP.Z / nP
    m3PsCenter = cP
End Function

Public Function m3VectorCross(ByRef v As Vector, ByRef u As Vector) As Vector
    Dim w As Vector
    w.x = v.y * u.Z - v.Z * u.y
    w.y = v.Z * u.x - v.x * u.Z
    w.Z = v.x * u.y - v.y * u.x
    m3VectorCross = w
End Function

Public Function m3LineRotate(ByRef p As m3Point, ByRef D1 As Vector, ByVal zavit As Double) As m3Matrix
    Dim m As m3Matrix
    Dim Cs As Double
    Dim Sn As Double
    Dim c As Double
    Dim d As Vector
    d = D1
    m3ops.m3VectorSetLength 1, d
    Cs = Cos(zavit)
    Sn = Sin(zavit)
    c = 1 - Cs
    m.mat(0, 0) = d.x * d.x * c + Cs
    m.mat(0, 1) = d.x * d.y * c + d.Z * Sn
    m.mat(0, 2) = d.x * d.Z * c - d.y * Sn
    m.mat(1, 0) = d.y * d.x * c - d.Z * Sn
    m.mat(1, 1) = d.y * d.y * c + Cs
    m.mat(1, 2) = d.y * d.Z * c + d.x * Sn
    m.mat(2, 0) = d.Z * d.x * c + d.y * Sn
    m.mat(2, 1) = d.Z * d.y * c - d.x * Sn
    m.mat(2, 2) = d.Z * d.Z * c + Cs
    m.mat(3, 0) = p.x - p.x * m.mat(0, 0) - p.y * m.mat(1, 0) - p.Z * m.mat(2, 0)
    m.mat(3, 1) = p.y - p.x * m.mat(0, 1) - p.y * m.mat(1, 1) - p.Z * m.mat(2, 1)
    m.mat(3, 2) = p.Z - p.x * m.mat(0, 2) - p.y * m.mat(1, 2) - p.Z * m.mat(2, 2)
    m3LineRotate = m
End Function

Public Function m3VectorDot(ByRef v As Vector, ByRef u As Vector) As Double
    m3VectorDot = v.x * u.x + v.y * u.y + v.Z * u.Z
End Function

Public Function m3inFrontOfPlane(ByRef n As Vector, ByRef p0 As m3Point, ByRef p As m3Point) As Boolean
    Dim B As Boolean
    Dim d As Double
    Dim v As Vector
    v = m3VectorInit(p0, p)
    m3inFrontOfPlane = m3ops.m3VectorDot(n, v) <= 0
End Function

Public Function m3VectorToY(ByRef p As m3Point, ByRef v As Vector) As m3Matrix
    Dim m As m3Matrix
    Dim l As Double
    Dim d As Vector
    d = v
    m3ops.m3VectorSetLength 1, d
    l = Sqr(d.x * d.x + d.y * d.y)
    If l <= 0.00000001 Then
        If d.Z < 0 Then
            m = m3ops.m3RotateX(Pie / 2)
            m = m3MatMultiply(m3Translate(-p.x, -p.y, -p.Z), m)
        Else
            m = m3ops.m3RotateX(-Pie / 2)
            m = m3MatMultiply(m3Translate(-p.x, -p.y, -p.Z), m)
        End If
        m3VectorToY = m
        Exit Function
    End If
    m.mat(0, 0) = d.y / l
    m.mat(0, 1) = d.x
    m.mat(0, 2) = -d.x * d.Z / l
    m.mat(1, 0) = -d.x / l
    m.mat(1, 1) = d.y
    m.mat(1, 2) = -d.y * d.Z / l
    m.mat(2, 0) = 0
    m.mat(2, 1) = d.Z
    m.mat(2, 2) = l
    m.mat(3, 0) = 0
    m.mat(3, 1) = 0
    m.mat(3, 2) = 0
    m = m3MatMultiply(m3Translate(-p.x, -p.y, -p.Z), m)
    m3VectorToY = m
End Function

Public Function sqrEquasion(ByVal a As Double, ByVal B As Double, ByVal c As Double) As Double()
    'פעולה זו פותרת משוואה ריבועית
    Dim d As Double
    Dim solution() As Double
    d = B * B - 4 * a * c
    If d < 0 Or a = 0 Then
        ReDim solution(0) As Double
        sqrEquasion = solution
        Exit Function
    Else
        If d = 0 Then
            ReDim solution(1) As Double
            solution(0) = -B / (2 * a)
            sqrEquasion = solution
            Exit Function
        End If
    End If
    ReDim solution(2) As Double
    solution(0) = (-B - Sqr(d)) / (2 * a)
    solution(1) = (-B + Sqr(d)) / (2 * a)
    sqrEquasion = solution
End Function
