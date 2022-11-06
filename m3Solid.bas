Attribute VB_Name = "m3Solid"
Option Explicit
Public Type Face3D
    nVerts As Single
    Face() As Integer
End Type
Public Type Solid
    nVerts As Single
    nFaces As Single
    verts() As m3Point
    Faces() As Face3D
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
Dim pt(1000) As POINTAPI
Dim pt_proj(10000) As m3Point
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Function m3SolidInit(ByVal Path As String) As Solid
    Dim s As Solid
    Dim FileNum As Single
    Dim i As Single
    Dim j As Single
    Dim k As Single
    FileNum = FreeFile
    Open Path For Input As #FileNum
    Input #FileNum, s.nVerts
    Input #FileNum, s.nFaces
    ReDim s.verts(s.nVerts - 1) As m3Point
    For i = 0 To s.nVerts - 1
        Input #FileNum, s.verts(i).x, s.verts(i).y, s.verts(i).Z
    Next i
    ReDim s.Faces(s.nFaces - 1) As Face3D
    For j = 0 To s.nFaces - 1
        Input #FileNum, s.Faces(j).nVerts
        ReDim s.Faces(j).Face(s.Faces(j).nVerts - 1) As Integer
        For k = 0 To s.Faces(j).nVerts - 1
            Input #FileNum, s.Faces(j).Face(k)
        Next k
    Next j
    Close #FileNum
    m3SolidInit = s
End Function

Public Sub m3SolidApply(ByRef s As Solid, ByRef m As m3Matrix)
    Dim i As Single
    For i = 0 To s.nVerts - 1
        m3Pointapply s.verts(i), m
    Next i
End Sub

Public Sub SolidDraw(ByRef obj As Object, ByRef s As Solid)
    Dim xOrg As Single
    Dim yOrg As Single
    Dim i As Single
    Dim j As Single
    Dim p As Single
    Dim P1 As Single
    Dim P2 As Single
    Dim P3 As Single
    Dim v As Vector
    Dim u As Vector
    Dim w As Vector
    Dim dist As Double
    Dim f As Double
    xOrg = Abs(obj.ScaleLeft)
    yOrg = Abs(obj.ScaleTop)
    If m3isparallel Then
        For i = 0 To s.nVerts - 1
            pt_proj(i) = s.verts(i)
        Next i
    Else
        dist = m3getdistance
        For i = 0 To s.nVerts - 1
            f = dist / (dist - s.verts(i).Z)
            pt_proj(i).x = s.verts(i).x * f
            pt_proj(i).y = s.verts(i).y * f
            pt_proj(i).Z = s.verts(i).Z
        Next i
    End If
    For i = 0 To s.nFaces - 1
        P1 = s.Faces(i).Face(0)
        P2 = s.Faces(i).Face(1)
        P3 = s.Faces(i).Face(2)
        v = m3ops.m3VectorInit(pt_proj(P1), pt_proj(P2))
        u = m3VectorInit(pt_proj(P2), pt_proj(P3))
        w = m3ops.m3VectorCross(v, u)
        If w.Z >= 0 Then
            For j = 0 To s.Faces(i).nVerts - 1
                p = s.Faces(i).Face(j)
                pt(j).x = xOrg + pt_proj(p).x
                pt(j).y = yOrg - pt_proj(p).y
            Next j
            Polygon obj.hdc, pt(0), s.Faces(i).nVerts
        End If
    Next i
End Sub
Public Sub SolidDrawShading(ByRef obj As Object, ByRef s As Solid)
    Dim xOrg As Single
    Dim yOrg As Single
    Dim i As Single
    Dim j As Single
    Dim p As Single
    Dim P1 As Single
    Dim P2 As Single
    Dim P3 As Single
    Dim v As Vector
    Dim u As Vector
    Dim w As Vector
    Dim dist As Double
    Dim f As Double
    Dim light As Vector
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim OldFillColor As Long
    Dim OldForeColor As Long
    Dim ColorTo As Long
    OldFillColor = obj.FillColor
    OldForeColor = obj.ForeColor
    Dim dot As Double
    xOrg = Abs(obj.ScaleLeft)
    yOrg = Abs(obj.ScaleTop)
    light = m3draw.GetLightVector
    R = OldFillColor And &HFF&
    G = (OldFillColor And &HFF00&) \ &H100&
    B = (OldFillColor And &HFF0000) \ &H10000
    If m3isparallel Then
        For i = 0 To s.nVerts - 1
            pt_proj(i) = s.verts(i)
        Next i
    Else
        dist = m3getdistance
        For i = 0 To s.nVerts - 1
            f = dist / (dist - s.verts(i).Z)
            pt_proj(i).x = s.verts(i).x * f
            pt_proj(i).y = s.verts(i).y * f
            pt_proj(i).Z = s.verts(i).Z
        Next i
    End If
    For i = 0 To s.nFaces - 1
        P1 = s.Faces(i).Face(0)
        P2 = s.Faces(i).Face(1)
        P3 = s.Faces(i).Face(2)
        v = m3ops.m3VectorInit(pt_proj(P1), pt_proj(P2))
        u = m3VectorInit(pt_proj(P2), pt_proj(P3))
        w = m3ops.m3VectorCross(v, u)
        If w.Z >= 0 Then
            
            For j = 0 To s.Faces(i).nVerts - 1
                p = s.Faces(i).Face(j)
                pt(j).x = xOrg + pt_proj(p).x
                pt(j).y = yOrg - pt_proj(p).y
            Next j
            m3ops.m3VectorSetLength 1, w
            dot = m3ops.m3VectorDot(w, light)
            If dot < 0.000001 Then
                dot = 0.3
            Else
                dot = dot + 0.3
                If dot > 1 Then dot = 1
            End If
            ColorTo = RGB(R * dot, G * dot, B * dot)
            obj.FillColor = ColorTo
            obj.ForeColor = ColorTo
            Polygon obj.hdc, pt(0), s.Faces(i).nVerts
        End If
    Next i
    obj.FillColor = OldFillColor
    obj.ForeColor = OldFillColor
End Sub
Public Sub SolidDrawShadingDot(ByRef obj As Object, ByRef s As Solid)
    Dim xOrg As Single
    Dim yOrg As Single
    Dim i As Single
    Dim j As Single
    Dim p As Single
    Dim P1 As Single
    Dim P2 As Single
    Dim P3 As Single
    Dim v As Vector
    Dim u As Vector
    Dim w As Vector
    Dim dist As Double
    Dim f As Double
    Dim light As Vector
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim c As m3Point
    Dim OldFillColor As Long
    Dim OldForeColor As Long
    Dim ColorTo As Long
    OldFillColor = obj.FillColor
    OldForeColor = obj.ForeColor
    Dim dot As Double
    xOrg = Abs(obj.ScaleLeft)
    yOrg = Abs(obj.ScaleTop)
    
    R = OldFillColor And &HFF&
    G = (OldFillColor And &HFF00&) \ &H100&
    B = (OldFillColor And &HFF0000) \ &H10000
    If m3isparallel Then
        For i = 0 To s.nVerts - 1
            pt_proj(i) = s.verts(i)
        Next i
    Else
        dist = m3getdistance
        
        For i = 0 To s.nVerts - 1
            f = dist / (dist - s.verts(i).Z)
            pt_proj(i).x = s.verts(i).x * f
            pt_proj(i).y = s.verts(i).y * f
            pt_proj(i).Z = s.verts(i).Z
        Next i
    End If
    
    For i = 0 To s.nFaces - 1
        P1 = s.Faces(i).Face(0)
        P2 = s.Faces(i).Face(1)
        P3 = s.Faces(i).Face(2)
        v = m3ops.m3VectorInit(pt_proj(P1), pt_proj(P2))
        u = m3VectorInit(pt_proj(P2), pt_proj(P3))
        w = m3ops.m3VectorCross(v, u)
        If w.Z >= 0 Then
            c.x = 0
            c.y = 0
            c.Z = 0
            For j = 0 To s.Faces(i).nVerts - 1
                p = s.Faces(i).Face(j)
                c.x = c.x + s.verts(p).x
                c.y = c.y + s.verts(p).y
                c.Z = c.Z + s.verts(p).Z
                pt(j).x = xOrg + pt_proj(p).x
                pt(j).y = yOrg - pt_proj(p).y
            Next j
            c.x = c.x / s.Faces(i).nVerts
            c.y = c.y / s.Faces(i).nVerts
            c.Z = c.Z / s.Faces(i).nVerts
            light = m3ops.m3VectorInit(c, m3PointInit(100, 200, 200))
            m3ops.m3VectorSetLength 1, light
            m3ops.m3VectorSetLength 1, w
            dot = m3ops.m3VectorDot(w, light)
            If dot < 0.000001 Then
                dot = 0.3
            Else
                dot = dot + 0.3
                If dot > 1 Then dot = 1
            End If
            ColorTo = RGB(R * dot, G * dot, B * dot)
            obj.FillColor = ColorTo
            obj.ForeColor = ColorTo
            Polygon obj.hdc, pt(0), s.Faces(i).nVerts
        End If
    Next i
    obj.FillColor = OldFillColor
    obj.ForeColor = OldFillColor
End Sub

Public Function m3SolidCenter(ByRef s As Solid) As m3Point
    Dim i As Single
    Dim cP As m3Point
    cP.x = 0: cP.y = 0: cP.Z = 0
    For i = 0 To s.nVerts - 1
        cP.x = cP.x + s.verts(i).x
        cP.y = cP.y + s.verts(i).y
        cP.Z = cP.Z + s.verts(i).Z
    Next i
    cP.x = cP.x / s.nVerts
    cP.y = cP.y / s.nVerts
    cP.Z = cP.Z / s.nVerts
    m3SolidCenter = cP
End Function

Public Function CilinderInit(ByVal R As Single, ByVal h As Single, ByVal ppBase As Single) As Solid
    Dim s As Solid
    Dim i As Single
    Dim p As m3Point
    Dim mY As m3Matrix
    s.nVerts = 2 * ppBase
    ReDim s.verts(s.nVerts - 1) As m3Point
    s.nFaces = ppBase + 2
    ReDim s.Faces(s.nFaces - 1) As Face3D
    p = m3PointInit(R, 0, 0)
    mY = m3RotateY(2 * Pie / ppBase)
    For i = 0 To ppBase - 1
        s.verts(i) = p
        s.verts(i + ppBase) = p
        s.verts(i + ppBase).y = h
        m3Pointapply p, mY
    Next i
    s.Faces(0).nVerts = ppBase
    s.Faces(1).nVerts = ppBase
    ReDim s.Faces(0).Face(ppBase) As Integer
    ReDim s.Faces(1).Face(ppBase) As Integer
    For i = 0 To ppBase - 1
        s.Faces(0).Face(i) = ppBase - i - 1
        s.Faces(1).Face(i) = i + ppBase
    Next i
    For i = 0 To ppBase - 1
        s.Faces(i + 2).nVerts = 4
        ReDim s.Faces(i + 2).Face(3) As Integer
        s.Faces(i + 2).Face(0) = i
        s.Faces(i + 2).Face(1) = (i + 1) Mod ppBase
        s.Faces(i + 2).Face(2) = (i + 1) Mod ppBase + ppBase
        s.Faces(i + 2).Face(3) = i + ppBase
    Next i
    CilinderInit = s
End Function

Public Function ConeInit(ByVal R As Single, ByVal h As Single, ByVal ppBase As Single) As Solid
    Dim s As Solid
    Dim i As Single
    Dim p As m3Point
    Dim mY As m3Matrix
    s.nVerts = ppBase + 1
    ReDim s.verts(s.nVerts - 1) As m3Point
    s.nFaces = ppBase + 1
    ReDim s.Faces(s.nFaces - 1) As Face3D
    p = m3PointInit(R, 0, 0)
    mY = m3RotateY(2 * Pie / ppBase)
    For i = 0 To ppBase
        s.verts(i) = p
        m3Pointapply p, mY
    Next i
    s.verts(s.nVerts - 1).x = 0: s.verts(s.nVerts - 1).y = h: s.verts(s.nVerts - 1).Z = 0
    s.Faces(0).nVerts = ppBase
    ReDim s.Faces(0).Face(ppBase) As Integer
    For i = 0 To ppBase - 1
        s.Faces(0).Face(i) = ppBase - i - 1
    Next i
    For i = 0 To ppBase - 1
        s.Faces(i + 1).nVerts = 3
        ReDim s.Faces(i + 1).Face(2) As Integer
        s.Faces(i + 1).Face(0) = i
        s.Faces(i + 1).Face(1) = (i + 1) Mod ppBase
        s.Faces(i + 1).Face(2) = ppBase
    Next i
    ConeInit = s
End Function

Public Function SphereInit(ByVal R As Double, ByVal ppSlice As Double) As Solid
    Dim s As Solid
    Dim nSlices As Single
    Dim mZ As m3Matrix
    Dim mY As m3Matrix
    Dim p As m3Point
    Dim i As Single
    Dim j As Single
    Dim ind As Single
    Dim p0 As Integer
    nSlices = (ppSlice - 2) \ 2
    s.nVerts = ppSlice * nSlices + 2
    ReDim s.verts(s.nVerts - 1) As m3Point
    mZ = m3RotateZ(2 * Pie / ppSlice)
    mY = m3RotateY(2 * Pie / ppSlice)
    p = m3PointInit(0, R, 0)
    s.verts(0) = p
    m3Pointapply p, mZ
    For i = 0 To nSlices - 1
        For j = 1 To ppSlice
            s.verts(i * ppSlice + j) = p
            m3Pointapply p, mY
        Next j
        m3Pointapply p, mZ
    Next i
    s.verts(s.nVerts - 1) = p
    s.nFaces = 2 * ppSlice + (nSlices - 1) * ppSlice
    ReDim s.Faces(s.nFaces) As Face3D
    For i = 0 To ppSlice - 1
        s.Faces(i).nVerts = 3
        ReDim s.Faces(i).Face(3) As Integer
        s.Faces(i).Face(0) = 0
        s.Faces(i).Face(1) = i + 1
        s.Faces(i).Face(2) = (i + 1) Mod ppSlice + 1
        ind = i + ppSlice
        s.Faces(i + ppSlice).nVerts = 3
        ReDim s.Faces(ind).Face(3) As Integer
        s.Faces(ind).Face(2) = s.nVerts - 2 - i
        s.Faces(ind).Face(1) = s.nVerts - 1
        s.Faces(ind).Face(0) = s.nVerts - 2 - (i + 1) Mod ppSlice
    Next i
    For i = 0 To nSlices - 2
        For j = 0 To ppSlice - 1
            ind = ppSlice * 2 + i * ppSlice + j
            s.Faces(ind).nVerts = 4
            ReDim s.Faces(ind).Face(4) As Integer
            p0 = 1 + i * ppSlice
            s.Faces(ind).Face(0) = p0 + j
            s.Faces(ind).Face(1) = p0 + ppSlice + j
            s.Faces(ind).Face(2) = p0 + (j + 1) Mod ppSlice + ppSlice
            s.Faces(ind).Face(3) = p0 + (j + 1) Mod ppSlice
        Next j
    Next i
    SphereInit = s
End Function

Public Function HalfCilinder(ByVal R As Integer, ByVal h As Integer, ByVal ppBase As Integer) As Solid
    Dim s As Solid
    Dim i As Single
    Dim p As m3Point
    Dim mY As m3Matrix
    s.nVerts = ppBase * 2
    ReDim s.verts(s.nVerts - 1) As m3Point
    s.nFaces = ppBase + 2
    ReDim s.Faces(s.nFaces - 1) As Face3D
    p = m3PointInit(R, 0, 0)
    mY = m3RotateY(Pie / (ppBase - 1))
    For i = 0 To ppBase - 1
        s.verts(i) = p
        s.verts(i + ppBase) = p
        s.verts(i + ppBase).y = h
        m3Pointapply p, mY
    Next i
    s.Faces(0).nVerts = ppBase
    s.Faces(1).nVerts = ppBase
    ReDim s.Faces(0).Face(ppBase) As Integer
    ReDim s.Faces(1).Face(ppBase) As Integer
    For i = 0 To ppBase - 1
        s.Faces(0).Face(i) = ppBase - i - 1
        s.Faces(1).Face(i) = i + ppBase
    Next i
    For i = 0 To ppBase - 1
        s.Faces(i + 2).nVerts = 4
        ReDim s.Faces(i + 2).Face(3) As Integer
        s.Faces(i + 2).Face(0) = i
        s.Faces(i + 2).Face(1) = (i + 1) Mod ppBase
        s.Faces(i + 2).Face(2) = (i + 1) Mod ppBase + ppBase
        s.Faces(i + 2).Face(3) = i + ppBase
    Next i
    HalfCilinder = s
End Function
Public Function GuitarNeck(ByVal R As Double, ByVal h As Double, ByVal ppBase As Integer) As Solid
    Dim s As Solid
    Dim i As Single
    Dim p As m3Point
    Dim mY As m3Matrix
    Dim mX As m3Matrix
    Dim m As m3Matrix
    s.nVerts = ppBase * 2
    ReDim s.verts(s.nVerts - 1) As m3Point
    s.nFaces = ppBase + 2
    ReDim s.Faces(s.nFaces - 1) As Face3D
    p = m3PointInit(R, 0, 0)
    mY = m3RotateY(Pie / (ppBase - 1))
    
    For i = 0 To ppBase - 1
        s.verts(i) = p
        s.verts(i + ppBase) = p
        s.verts(i + ppBase).y = h
        m3Pointapply p, mY
    Next i
    s.Faces(0).nVerts = ppBase
    s.Faces(1).nVerts = ppBase
    ReDim s.Faces(0).Face(ppBase) As Integer
    ReDim s.Faces(1).Face(ppBase) As Integer
    For i = 0 To ppBase - 1
        s.Faces(0).Face(i) = ppBase - i - 1
        s.Faces(1).Face(i) = i + ppBase
    Next i
    For i = 0 To ppBase - 1
        s.Faces(i + 2).nVerts = 4
        ReDim s.Faces(i + 2).Face(3) As Integer
        s.Faces(i + 2).Face(0) = i
        s.Faces(i + 2).Face(1) = (i + 1) Mod ppBase
        s.Faces(i + 2).Face(2) = (i + 1) Mod ppBase + ppBase
        s.Faces(i + 2).Face(3) = i + ppBase
    Next i
    
    mX = m3Scale(2, 1, 1)
    m3SolidApply s, mX
    GuitarNeck = s
End Function

Public Function m3Box(ByVal d As Double, ByVal w As Double, ByVal h As Double) As Solid
    Dim s As Solid
    Dim i As Integer
    Dim num As Integer
    s.nVerts = 8
    s.nFaces = 6
    ReDim s.verts(s.nVerts - 1) As m3Point
    ReDim s.Faces(s.nFaces - 1) As Face3D
    'front face
    s.verts(0) = m3ops.m3PointInit(-w / 2, -h / 2, d / 2)
    s.verts(1) = m3ops.m3PointInit(w / 2, -h / 2, d / 2)
    s.verts(2) = m3ops.m3PointInit(w / 2, h / 2, d / 2)
    s.verts(3) = m3ops.m3PointInit(-w / 2, h / 2, d / 2)
    'back face
    s.verts(4) = m3ops.m3PointInit(-w / 2, -h / 2, -d / 2)
    s.verts(5) = m3ops.m3PointInit(w / 2, -h / 2, -d / 2)
    s.verts(6) = m3ops.m3PointInit(w / 2, h / 2, -d / 2)
    s.verts(7) = m3ops.m3PointInit(-w / 2, h / 2, -d / 2)
    
    'front face
    
    s.Faces(0).nVerts = 4
    ReDim s.Faces(0).Face(s.Faces(0).nVerts - 1)
    For i = 0 To 3
        s.Faces(0).Face(i) = i
    Next i
    
    'back face
    
    s.Faces(1).nVerts = 4
    ReDim s.Faces(1).Face(s.Faces(1).nVerts - 1)
    
    For i = 3 To 0 Step -1
        s.Faces(1).Face(-1 * i + 3) = i + 4
    Next i
   
    
    'Right face
    s.Faces(2).nVerts = 4
    ReDim s.Faces(2).Face(s.Faces(2).nVerts - 1)

    s.Faces(2).Face(0) = 5
    s.Faces(2).Face(1) = 6
    s.Faces(2).Face(2) = 2
    s.Faces(2).Face(3) = 1
    
        
    'left face
    s.Faces(3).nVerts = 4
    ReDim s.Faces(3).Face(s.Faces(3).nVerts - 1)

    s.Faces(3).Face(0) = 3
    s.Faces(3).Face(1) = 7
    s.Faces(3).Face(2) = 4
    s.Faces(3).Face(3) = 0
    
    'up face
        s.Faces(4).nVerts = 4
    ReDim s.Faces(4).Face(s.Faces(4).nVerts - 1)

    s.Faces(4).Face(0) = 2
    s.Faces(4).Face(1) = 6
    s.Faces(4).Face(2) = 7
    s.Faces(4).Face(3) = 3
    
    'buttom face
        s.Faces(5).nVerts = 4
    ReDim s.Faces(5).Face(s.Faces(5).nVerts - 1)

    s.Faces(5).Face(0) = 0
    s.Faces(5).Face(1) = 4
    s.Faces(5).Face(2) = 5
    s.Faces(5).Face(3) = 1
    
    m3Box = s
End Function

Public Function m3Mesholash(ByVal d As Double, ByVal h As Double, ByVal w As Double) As Solid
    Dim m As Solid
    m.nVerts = 6
    ReDim m.verts(m.nVerts - 1) As m3Point
    m.nFaces = 5
    ReDim m.Faces(m.nFaces - 1) As Face3D
    m.verts(5).x = 0: m.verts(5).y = 0: m.verts(5).Z = 0
    m.verts(4).x = w: m.verts(4).y = 0: m.verts(4).Z = 0
    m.verts(3).x = w: m.verts(3).y = h: m.verts(3).Z = 0
    m.verts(2).x = w: m.verts(2).y = h: m.verts(2).Z = d
    m.verts(1).x = w: m.verts(1).y = 0: m.verts(1).Z = d
    m.verts(0).x = 0: m.verts(0).y = 0: m.verts(0).Z = d
    
    'FrontFace
    m.Faces(0).nVerts = 3
    ReDim m.Faces(0).Face(2)
    m.Faces(0).Face(0) = 0
    m.Faces(0).Face(1) = 1
    m.Faces(0).Face(2) = 2
    
    'BackFace
    m.Faces(1).nVerts = 3
    ReDim m.Faces(1).Face(2)
    m.Faces(1).Face(0) = 3
    m.Faces(1).Face(1) = 4
    m.Faces(1).Face(2) = 5
    
    'BottomFace
    m.Faces(2).nVerts = 4
    ReDim m.Faces(2).Face(3)
    m.Faces(2).Face(0) = 0
    m.Faces(2).Face(1) = 5
    m.Faces(2).Face(2) = 4
    m.Faces(2).Face(3) = 1
    
    'SideStraight
    m.Faces(3).nVerts = 4
    ReDim m.Faces(3).Face(3)
    m.Faces(3).Face(0) = 1
    m.Faces(3).Face(1) = 4
    m.Faces(3).Face(2) = 3
    m.Faces(3).Face(3) = 2
    
    'Side
    m.Faces(4).nVerts = 4
    ReDim m.Faces(4).Face(3)
    m.Faces(4).Face(0) = 2
    m.Faces(4).Face(1) = 3
    m.Faces(4).Face(2) = 5
    m.Faces(4).Face(3) = 0
    m3Mesholash = m
End Function

