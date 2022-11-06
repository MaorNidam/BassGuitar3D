Attribute VB_Name = "m3Guitar"
Option Explicit
Private Const NoCollision = -100000
Private Const ppBase = 6
Private Const R = 50

Public Type m3StringType
        nLinks As Integer
        Links() As Solid
        R As Double
        l As Double
End Type
Public Type Strings
    nStrings As Integer
    Strings() As m3StringType
    Color As Long
End Type
Public Type Guitar
    Body1(3) As Solid
    nBody1S As Integer
    cBody As Long
    Body2(3) As Solid
    nBody2S As Integer
    Head(4) As Solid
    cHead As Long
    nHeadS As Integer
    Neck As Solid
    cNeck As Long
    cNeckC As Long
    cPickup As Long
    Wall(2) As Solid
    Pickup(4) As Solid
    cWallG As Long
    cWallH As Long
    nWalls As Integer
    nPickups As Integer
End Type

Public Type Note
    Scale As String
    Note As String
End Type
Public Function GuitarInit(ByRef s As Strings) As Guitar
    '.פעולה זו מהתחלת את הנתונים של הגיטרה על ידי רדיוס חצי הגליל בגוף הגיטרה
    Dim G As Guitar
    Dim v As Vector
    Dim u As Vector
    Dim i As Integer
    Dim m As m3Matrix
    Dim l As Double
    Dim k As Double
    Dim st As Strings
    Dim Cs As Double
    Dim Sn As Double
    Dim alfa As Double
    Dim p As m3Point
    G.nBody1S = 4
    G.nBody2S = 4
    G.nHeadS = 5
    G.nWalls = 2
    G.nPickups = 5
    G.cBody = RGB(189, 151, 79)
    G.cHead = RGB(189, 151, 79)
    G.cNeck = RGB(105, 86, 61)
    G.cNeckC = RGB(74, 61, 43)
    G.cPickup = vbBlack
    G.cWallG = RGB(192, 192, 210)
    G.cWallH = vbWhite
    
   '*******g.Body1(0) = Round body part**************************************
    G.Body1(0) = HalfCilinder(R, 10, 16)
    G.Body1(0).nFaces = G.Body1(0).nFaces - 1
    '*******g.Body1(1) = box body part**************************************
    G.Body1(1) = m3Box(R / 2, 10, 2 * R)
    m3SolidApply G.Body1(1), m3ops.m3Translate(-R, 0, 0)
    m3SolidApply G.Body1(1), m3RotateZ(Pie / 2)
    v = m3VectorInit(G.Body1(1).verts(6), G.Body1(0).verts(0))
    m3SolidApply G.Body1(1), m3Translate(v.x, v.y, v.Z)
    
    
    For i = 2 To G.Body1(1).nFaces - 1
        G.Body1(1).Faces(i - 2) = G.Body1(1).Faces(i)
    Next i
    G.Body1(1).nFaces = G.Body1(1).nFaces - 2
    '*******g.Body1(2) = Trapez body part**************************************
    G.Body1(2) = HalfCilinder(R, 10, 4)
    m3SolidApply G.Body1(2), m3RotateX(Pie)
    v = m3VectorInit(G.Body1(2).verts(3), G.Body1(1).verts(0))
    m3SolidApply G.Body1(2), m3Translate(v.x, v.y, v.Z)
    'Fitting smaller base to g.neck
    m3Pointapply G.Body1(2).verts(1), m3Translate(-R / 5, 0, 0)
    m3Pointapply G.Body1(2).verts(2), m3Translate(R / 5, 0, 0)
    m3Pointapply G.Body1(2).verts(5), m3Translate(-R / 5, 0, 0)
    m3Pointapply G.Body1(2).verts(6), m3Translate(R / 5, 0, 0)
    
    G.Body1(2).nFaces = G.Body1(2).nFaces - 1
    '*******g.Body2(0) = Right triangle**************************************
    'g.Body2(0) = m3Mesholash(10, h / 3, R)
    G.Body2(0) = m3Mesholash(10, 2 * R / 3, R)
    m3SolidApply G.Body2(0), m3RotateX(-Pie / 2)
    'Even with trapez
    v = m3VectorInit(G.Body1(2).verts(3), G.Body1(2).verts(2))
    m3ops.m3VectorSetLength 1, v
    m = m3Identity
    m.mat(0, 0) = v.Z
    m.mat(0, 2) = -v.x
    m.mat(2, 0) = v.x
    m.mat(2, 2) = v.Z
    m3SolidApply G.Body2(0), m
    v = m3VectorInit(G.Body2(0).verts(1), G.Body1(2).verts(2))
    m3SolidApply G.Body2(0), m3Translate(v.x, v.y, v.Z)
    'move to place
    v = m3VectorInit(G.Body2(0).verts(1), G.Body1(2).verts(2))
    m3SolidApply G.Body2(0), m3Translate(v.x, v.y, v.Z)
    
    G.Body2(0).Faces(G.Body2(0).nFaces - 2) = G.Body2(0).Faces(G.Body2(0).nFaces - 1)
    G.Body2(0).nFaces = G.Body2(0).nFaces - 1
    '*******g.Body2(1) = Left triangle**************************************
    G.Body2(1) = m3Mesholash(10, 2 * R / 3, R)
    m3SolidApply G.Body2(1), m3RotateX(Pie / 2)
    'even with trapez
    v = m3VectorInit(G.Body1(2).verts(1), G.Body1(2).verts(0))
    m3ops.m3VectorSetLength 1, v
    m = m3Identity
    m.mat(0, 0) = v.Z
    m.mat(0, 2) = -v.x
    m.mat(2, 0) = v.x
    m.mat(2, 2) = v.Z
    m3SolidApply G.Body2(1), m
    'move to place
    v = m3VectorInit(G.Body2(1).verts(4), G.Body1(2).verts(1))
    m3SolidApply G.Body2(1), m3Translate(v.x, v.y, v.Z)
    
    G.Body2(1).Faces(G.Body2(1).nFaces - 2) = G.Body2(1).Faces(G.Body2(1).nFaces - 1)
    G.Body2(1).nFaces = G.Body2(1).nFaces - 1
    '************g.neck = g.neck**************************************
    'get R
    v = m3VectorInit(G.Body1(2).verts(2), G.Body1(2).verts(1))
    G.Neck = GuitarNeck(m3VectorLength(v) / 4, R * 4, 16)
    m3SolidApply G.Neck, m3RotateX(-Pie / 2)
    'move to place
    v = m3VectorInit(G.Neck.verts(G.Neck.nVerts / 2), G.Body2(1).verts(4))
    m3SolidApply G.Neck, m3Translate(v.x, v.y, v.Z)
    For i = 0 To G.Neck.nVerts / 2 - 1
        m3Pointapply G.Neck.verts(i), m3Scale(0.8, 1, 1)
    Next i
    For i = 2 To G.Neck.nFaces - 1
        G.Neck.Faces(i - 2) = G.Neck.Faces(i)
    Next i
    G.Neck.nFaces = G.Neck.nFaces - 2
    '******************g.pickup(3)= Wood cover above g.neck***************************
    G.Pickup(3) = m3Box(R * 4 + R / 10, 3 * R / 5, 1)
    m3Pointapply G.Pickup(3).verts(0), m3Scale(0.8, 1, 1)
    m3Pointapply G.Pickup(3).verts(1), m3Scale(0.8, 1, 1)
    m3Pointapply G.Pickup(3).verts(2), m3Scale(0.8, 1, 1)
    m3Pointapply G.Pickup(3).verts(3), m3Scale(0.8, 1, 1)
    v = m3VectorInit(G.Pickup(3).verts(1), G.Neck.verts(0))
    m3SolidApply G.Pickup(3), m3Translate(v.x, v.y, v.Z)
    '********************Head part**************************************
    '**************g.head(0) = trapez ********************************
    G.Head(0) = HalfCilinder(3 * R / 5, 10, 4)
    G.Head(0).nFaces = G.Head(0).nFaces - 1
    For i = 0 To G.Head(0).nVerts - 1
        m3Pointapply G.Head(0).verts(i), m3Scale(0.75, 1, 1)
    Next i
    'move to place
    v = m3VectorInit(G.Head(0).verts(5), G.Neck.verts(0))
    m3SolidApply G.Head(0), m3Translate(v.x, v.y, v.Z)
    'adjust
    m3Pointapply G.Head(0).verts(0), m3Translate(0, 0, -15)
    m3Pointapply G.Head(0).verts(3), m3Translate(0, 0, -15)
    m3Pointapply G.Head(0).verts(4), m3Translate(0, 0, -15)
    m3Pointapply G.Head(0).verts(7), m3Translate(0, 0, -15)
    'L=trapez's big base length:
    v = m3VectorInit(G.Head(0).verts(4), G.Head(0).verts(7))
    l = m3VectorLength(v)
    '**************g.head(1) = Top triangle ********************************
    G.Head(1) = m3Mesholash(10, 40, l / 3)
    m3SolidApply G.Head(1), m3RotateY(Pie)
    m3SolidApply G.Head(1), m3RotateX(Pie / 2)
    v = m3VectorInit(G.Head(1).verts(0), G.Head(0).verts(4))
    m3SolidApply G.Head(1), m3Translate(v.x, v.y, v.Z)
    'g.Head(1).nFaces = g.Head(1).nFaces - 2
    'g.Head(1).nFaces = g.Head(1).nFaces - 3
    G.Head(1).Faces(G.Head(1).nFaces - 3) = G.Head(1).Faces(G.Head(1).nFaces - 1)
    G.Head(1).nFaces = G.Head(1).nFaces - 2
    '**************g.head(2) = Middle box ********************************
    G.Head(2) = m3Box(10, l / 3, 40)
    G.Head(2).Faces(G.Head(2).nFaces - 4) = G.Head(2).Faces(G.Head(2).nFaces - 2)
    G.Head(2).nFaces = G.Head(2).nFaces - 3
    m3SolidApply G.Head(2), m3RotateX(Pie / 2)
    v = m3VectorInit(G.Head(2).verts(1), G.Head(1).verts(4))
    m3SolidApply G.Head(2), m3Translate(v.x, v.y, v.Z)
    '*********g.head(3) = Bottom triangle***********************************
    G.Head(3) = G.Head(1)
    m3SolidApply G.Head(3), m3RotateZ(Pie)
    v = m3VectorInit(G.Head(3).verts(5), G.Head(0).verts(7))
    m3SolidApply G.Head(3), m3Translate(v.x, v.y, v.Z)
    
    G.Body1(2).verts(5) = G.Body2(1).verts(2)
    G.Body1(2).verts(1) = G.Body2(1).verts(3)
    G.Body1(2).verts(2) = G.Body2(0).verts(2)
    G.Body1(2).verts(6) = G.Body2(0).verts(3)
    v = m3VectorInit(G.Body2(0).verts(2), G.Body2(0).verts(0))
    u = m3VectorInit(G.Body2(0).verts(2), G.Body2(0).verts(1))
    k = u.Z / v.Z
    m3Pointapply G.Body2(0).verts(2), m3Translate(v.x * k, v.y * k, v.Z * k)
    m3Pointapply G.Body2(0).verts(3), m3Translate(v.x * k, v.y * k, v.Z * k)
    v = m3VectorInit(G.Body2(1).verts(2), G.Body2(1).verts(0))
    u = m3VectorInit(G.Body2(1).verts(2), G.Body2(1).verts(1))
    k = u.Z / v.Z
    m3Pointapply G.Body2(1).verts(2), m3Translate(v.x * k, v.y * k, v.Z * k)
    m3Pointapply G.Body2(1).verts(3), m3Translate(v.x * k, v.y * k, v.Z * k)
    G.Body2(2) = G.Body1(2)
    G.Body2(2).verts(6) = G.Body2(0).verts(3)
    G.Body2(2).verts(2) = G.Body2(0).verts(2)
    G.Body2(2).verts(7) = G.Body1(2).verts(6)
    G.Body2(2).verts(3) = G.Body1(2).verts(2)
    G.Body2(2).verts(0) = G.Body1(2).verts(1)
    G.Body2(2).verts(4) = G.Body1(2).verts(5)
    G.Body2(2).verts(1) = G.Body2(1).verts(3)
    G.Body2(2).verts(5) = G.Body2(1).verts(2)
    G.Pickup(0) = m3Box(2 * R / 5, R / 1.5, R / 15)
    G.Pickup(1) = m3Box(R / 5, R / 1.5, R / 15)
    G.Pickup(2) = m3Box(R / 5, R / 1.5, R / 15)
    m3SolidApply G.Pickup(1), m3Translate(0, 0, 1 * R / 5)
    m3SolidApply G.Pickup(2), m3Translate(0, 0, 4 * R / 5)
    G.Pickup(4) = m3Box(R / 10, 3 * R / 5, R / 15)
    G.Wall(0) = m3Box(6 * R / 50, R / 1.5, R / 15)
    v = m3VectorInit(G.Wall(0).verts(1), G.Pickup(0).verts(2))
    m3SolidApply G.Wall(0), m3Translate(v.x, v.y, v.Z)
    G.Wall(1) = m3Box(R / 10, 3 * R / 5, R / 15)
    v = m3VectorInit(G.Wall(1).verts(1), G.Pickup(4).verts(2))
    m3SolidApply G.Wall(1), m3Translate(v.x, v.y, v.Z)
    m3StringsApply s, m3RotateX(Pie / 2)
    m3StringsApply s, m3Translate(0, 0.3 * R, 0)
    v = m3VectorInit(s.Strings(0).Links(s.Strings(0).nLinks - 1).verts(s.Strings(0).Links(s.Strings(0).nLinks - 1).nVerts - 1), G.Pickup(3).verts(2))
    m3StringsApply s, m3Translate(v.x, v.y / 2, v.Z)
    st = s
    v = m3VectorInit(m3StringHeadCenter(s.Strings(0)), m3StringGuitarCenter(s.Strings(0)))
    u = m3VectorInit(G.Neck.verts(0), G.Neck.verts(G.Neck.nVerts / 2))
    Cs = m3ops.m3VectorDot(v, u) / (m3VectorLength(v) * m3VectorLength(u))
    Sn = Sqr(1 - Cs * Cs)
    alfa = Math.Atn(Sn / Cs)
    m3StringApply s.Strings(0), m3RotateY(-alfa)
    m3StringApply s.Strings(1), m3RotateY(-alfa * 0.8 / 2)
    m3StringApply s.Strings(2), m3RotateY(alfa * 0.8 / 2)
    m3StringApply s.Strings(3), m3RotateY(alfa)
    p.x = (G.Wall(0).verts(1).x + G.Wall(0).verts(2).x) / 2
    p.y = (G.Wall(0).verts(1).y + G.Wall(0).verts(2).y) / 2
    p.Z = (G.Wall(0).verts(1).Z + G.Wall(0).verts(2).Z) / 2
    v = m3VectorInit(p, m3StringGuitarCenter(s.Strings(0)))
    m3SolidApply G.Wall(0), m3Translate(0, v.y, v.Z)
    v = m3VectorInit(G.Pickup(0).verts(6), G.Wall(0).verts(5))
    m3SolidApply G.Pickup(0), m3Translate(v.x, v.y, v.Z)
    p.x = (G.Wall(1).verts(5).x + G.Wall(1).verts(6).x) / 2
    p.y = (G.Wall(1).verts(5).y + G.Wall(1).verts(6).y) / 2
    p.Z = (G.Wall(1).verts(5).Z + G.Wall(1).verts(6).Z) / 2
    v = m3VectorInit(p, m3StringHeadCenter(s.Strings(0)))
    m3SolidApply G.Wall(1), m3Translate(0, v.y, v.Z)
    v = m3VectorInit(G.Pickup(4).verts(6), G.Wall(1).verts(5))
    m3SolidApply G.Pickup(4), m3Translate(0, v.y, v.Z)
    v = m3VectorInit(G.Pickup(2).verts(3), G.Pickup(0).verts(3))
    m3SolidApply G.Pickup(2), m3Translate(0, v.y, 0)
    m3SolidApply G.Pickup(1), m3Translate(0, v.y, 0)
    
    GuitarApply G, m3RotateX(Pie / 2)
    GuitarApply G, m3RotateZ(-Pie / 2)
    m3StringsApply s, m3RotateX(Pie / 2)
    m3StringsApply s, m3RotateZ(-Pie / 2)
    GuitarApply G, m3Translate(-100, 0, 0)
    m3StringsApply s, m3Translate(-100, 0, 0)
    GuitarInit = G
End Function

Public Sub GuitarApply(ByRef G As Guitar, ByRef m As m3Matrix)
    Dim i As Integer
    For i = 0 To G.nBody1S - 1
        m3SolidApply G.Body1(i), m
    Next i
    For i = 0 To G.nBody2S - 1
        m3SolidApply G.Body2(i), m
    Next i
    For i = 0 To G.nHeadS - 1
        m3SolidApply G.Head(i), m
    Next i
    m3SolidApply G.Neck, m
    For i = 0 To G.nPickups - 1
        m3SolidApply G.Pickup(i), m
    Next i
    For i = 0 To G.nWalls - 1
        m3SolidApply G.Wall(i), m
    Next i
End Sub

Private Sub GuitarDrawTrianglesWithNec(ByRef obj As Object, ByRef G As Guitar)
    Dim left As Boolean ' body2(1)
    Dim right As Boolean ' body2(0)
    Dim oldc As Long
    left = m3PlaneIsVisible(G.Body2(0).verts(4), G.Body2(0).verts(1), G.Body2(0).verts(0))
    right = m3PlaneIsVisible(G.Body2(1).verts(4), G.Body2(1).verts(1), G.Body2(1).verts(0))
    oldc = obj.FillColor
    If left And right Then
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body2(0)
        m3Solid.SolidDrawShading obj, G.Body2(1)
        obj.FillColor = G.cNeck
        SolidDrawShading obj, G.Neck
        obj.FillColor = oldc
        Exit Sub
    End If
    If left Then
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body2(0)
        obj.FillColor = G.cNeck
        SolidDrawShading obj, G.Neck
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body2(1)
        obj.FillColor = oldc
        Exit Sub
    End If
    If right Then
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body2(1)
        obj.FillColor = G.cNeck
        SolidDrawShading obj, G.Neck
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body2(0)
        obj.FillColor = oldc
        Exit Sub
    End If
    obj.FillColor = G.cNeck
    SolidDrawShading obj, G.Neck
    obj.FillColor = G.cBody
    m3Solid.SolidDrawShading obj, G.Body2(0)
    m3Solid.SolidDrawShading obj, G.Body2(1)
    obj.FillColor = oldc
End Sub

Private Sub GuitarDrawTrianglesWithNecBody2_2(ByRef obj As Object, ByRef G As Guitar)
    Dim oldc As Long
    oldc = obj.FillColor
    If m3PlaneIsVisible(G.Body2(2).verts(1), G.Body2(2).verts(2), G.Body2(2).verts(6)) Then
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body2(2)
        GuitarDrawTrianglesWithNec obj, G
    Else
       GuitarDrawTrianglesWithNec obj, G
       obj.FillColor = G.cBody
       m3Solid.SolidDrawShading obj, G.Body2(2)
    End If
    obj.FillColor = oldc
End Sub

Private Sub GuitarDrawBody1TrianglesWithNecBody2_2(ByRef obj As Object, ByRef G As Guitar)
    Dim oldc As Long
    oldc = obj.FillColor
    If m3PlaneIsVisible(G.Body2(2).verts(3), G.Body2(2).verts(7), G.Body2(2).verts(4)) Then
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body1(0)
        m3Solid.SolidDrawShading obj, G.Body1(1)
        m3Solid.SolidDrawShading obj, G.Body1(2)
        GuitarDrawTrianglesWithNecBody2_2 obj, G
    Else
        GuitarDrawTrianglesWithNecBody2_2 obj, G
        obj.FillColor = G.cBody
        m3Solid.SolidDrawShading obj, G.Body1(0)
        m3Solid.SolidDrawShading obj, G.Body1(1)
        m3Solid.SolidDrawShading obj, G.Body1(2)
    End If
    obj.FillColor = oldc
End Sub

Private Sub GuitarDrawHeadBody1TrianglesWithNecBody2_2(ByRef obj As Object, ByRef G As Guitar)
    Dim oldc As Long
    If m3PlaneIsVisible(G.Head(0).verts(2), G.Head(0).verts(1), G.Head(0).verts(5)) Then
        GuitarDrawBody1TrianglesWithNecBody2_2 obj, G
        obj.FillColor = G.cHead
        m3Solid.SolidDrawShading obj, G.Head(1)
        m3Solid.SolidDrawShading obj, G.Head(2)
        m3Solid.SolidDrawShading obj, G.Head(3)
        m3Solid.SolidDrawShading obj, G.Head(0)
    Else
        obj.FillColor = G.cHead
        m3Solid.SolidDrawShading obj, G.Head(1)
        m3Solid.SolidDrawShading obj, G.Head(2)
        m3Solid.SolidDrawShading obj, G.Head(3)
        m3Solid.SolidDrawShading obj, G.Head(0)
        GuitarDrawBody1TrianglesWithNecBody2_2 obj, G
    End If
    obj.FillColor = oldc
End Sub

Public Sub GuitarDraw(ByRef obj As Object, ByRef G As Guitar, ByRef s As Strings)
    Dim oldc As Long
    oldc = obj.FillColor
    If m3PlaneIsVisible(G.Body2(2).verts(0), G.Body2(2).verts(1), G.Body2(2).verts(2)) Then
        DWallsWPickup obj, G, s
        GuitarDrawHeadBody1TrianglesWithNecBody2_2 obj, G
    Else
        GuitarDrawHeadBody1TrianglesWithNecBody2_2 obj, G
        DWallsWPickup obj, G, s
    End If
    obj.FillColor = oldc
End Sub

Public Sub PickupDraw(ByRef obj As Object, ByRef G As Guitar)
    Dim i As Integer
    Dim mid As Integer
    Dim P1 As m3Point
    Dim P2 As m3Point
    Dim P3 As m3Point
    Dim n As Integer
    Dim oldc As Long
    oldc = obj.FillColor
    n = G.nPickups - 1
    mid = n
    For i = 0 To G.nPickups - 1
        P1 = G.Pickup(i).verts(0)
        P2 = G.Pickup(i).verts(1)
        P3 = G.Pickup(i).verts(2)
        If Not m3draw.m3PlaneIsVisible(P1, P2, P3) Then
            mid = i - 1
            Exit For
        End If
    Next i
    For i = 0 To mid
        Select Case i
            Case 0
                obj.FillColor = G.cWallG
            Case 1, 2
                obj.FillColor = G.cPickup
            Case 3
                obj.FillColor = G.cNeckC
            Case 4
                obj.FillColor = G.cWallH
        End Select
        m3Solid.SolidDrawShading obj, G.Pickup(i)
    Next i
    For i = G.nPickups - 1 To mid + 1 Step -1
        Select Case i
            Case 0
                obj.FillColor = G.cWallG
            Case 1, 2
                obj.FillColor = G.cPickup
             Case 3
                obj.FillColor = G.cNeckC
            Case 4
                obj.FillColor = G.cWallH
            End Select
        m3Solid.SolidDrawShading obj, G.Pickup(i)
    Next i
    obj.FillColor = oldc
End Sub

Public Function m3StringInit(ByVal R As Double, ByVal l As Double, ByVal nLinks As Integer) As m3StringType
    Dim s As m3StringType
    Dim Sil As Solid
    Dim i As Integer
    Dim T As m3Matrix
    s.R = R
    s.l = l
    T = m3Translate(0, l / nLinks, 0)
    s.nLinks = nLinks
    ReDim s.Links(s.nLinks - 1) As Solid
    Sil = m3Solid.CilinderInit(R, l / nLinks, ppBase)
    For i = 0 To s.nLinks - 1
        s.Links(i) = Sil
        m3Solid.m3SolidApply Sil, T
    Next i
    m3StringInit = s
End Function

Public Sub m3StringApply(ByRef s As m3StringType, ByRef m As m3Matrix)
    Dim i As Integer
    For i = 0 To s.nLinks - 1
        m3Solid.m3SolidApply s.Links(i), m
    Next i
End Sub

Public Sub m3StringDraw(ByRef obj As Object, ByRef s As m3StringType)
    Dim i As Integer
    Dim midLink As Integer
    Dim P1 As m3Point
    Dim P2 As m3Point
    Dim P3 As m3Point
    Dim oldc As Long
    midLink = s.nLinks - 1
    
    oldc = obj.ForeColor
    For i = 0 To s.nLinks - 2
        P1 = s.Links(i).verts(s.Links(i).Faces(1).Face(0))
        P2 = s.Links(i).verts(s.Links(i).Faces(1).Face(1))
        P3 = s.Links(i).verts(s.Links(i).Faces(1).Face(2))
        If Not m3draw.m3PlaneIsVisible(P1, P2, P3) Then
            midLink = i - 1
            Exit For
        End If
    Next i
    obj.ForeColor = RGB(130, 130, 130)
    For i = 0 To midLink
        m3Solid.SolidDraw obj, s.Links(i)
    Next i
    For i = s.nLinks - 1 To midLink + 1 Step -1
        m3Solid.SolidDraw obj, s.Links(i)
    Next i
    obj.ForeColor = oldc
End Sub

Public Function m3StringCenter(ByRef s As m3StringType) As m3Point
    Dim cP As m3Point
    Dim Sp() As m3Point
    Dim i As Integer
    ReDim Sp(s.nLinks - 1) As m3Point
    cP.x = 0: cP.y = 0: cP.Z = 0
    For i = 0 To s.nLinks - 1
        Sp(i).x = 0: Sp(i).y = 0: Sp(i).Z = 0
    Next i
    For i = 0 To s.nLinks - 1
        Sp(i) = m3SolidCenter(s.Links(i))
    Next i
    For i = 0 To s.nLinks - 1
        cP.x = cP.x + Sp(i).x
        cP.y = cP.y + Sp(i).y
        cP.Z = cP.Z + Sp(i).Z
    Next i
    cP.x = cP.x / s.nLinks
    cP.y = cP.y / s.nLinks
    cP.Z = cP.Z / s.nLinks
    m3StringCenter = cP
End Function

Public Sub m3StringSwing(ByRef s As m3StringType, ByVal amp As Double, ByVal nLink As Integer)
    'פעולה זו מקבלת את המיתר הרצוי לתנועה, גודל האמפליטודה ואת החוליה במיתר הרצוי.
    'פעולה זו קובעת את כיוון תזוזת החלק במיתר בעת פריטתו
    Dim v As Vector
    Dim m As m3Matrix
    If Abs(amp) < 0 Then Exit Sub
    v = m3VectorInit(s.Links(0).verts(0), s.Links(0).verts(s.Links(0).nVerts \ 4))
    m3ops.m3VectorSetLength amp, v
    m = m3Translate(v.x, v.y, v.Z)
    m3Solid.m3SolidApply s.Links(nLink), m
End Sub

Public Function m3StringGuitarCenter(ByRef st As m3StringType) As m3Point
    'פעולה זו מחזירה את המרכז של הפאה הכי קיצונית של המיתר בצד הגוף של הגיטרה.
    Dim c As m3Point
    Dim i As Integer
    Dim n As Integer
    c.x = 0
    c.y = 0
    c.Z = 0
    n = st.Links(0).nVerts / 2 - 1
    For i = 0 To n
        c.x = c.x + st.Links(0).verts(i).x
        c.y = c.y + st.Links(0).verts(i).y
        c.Z = c.Z + st.Links(0).verts(i).Z
    Next i
    c.x = c.x / (n + 1)
    c.y = c.y / (n + 1)
    c.Z = c.Z / (n + 1)
    m3StringGuitarCenter = c
End Function
Public Function m3StringHeadCenter(ByRef st As m3StringType) As m3Point
    'פעולה זו מחזירה את המרכז של הפאה הכי קיצונית של המיתר בצד הראש של הגיטרה.
    Dim c As m3Point
    Dim i As Integer
    Dim n As Integer
    c.x = 0
    c.y = 0
    c.Z = 0
    n = st.Links(0).nVerts / 2 - 1
    For i = 0 To n
        c.x = c.x + st.Links(st.nLinks - 1).verts(i + n + 1).x
        c.y = c.y + st.Links(st.nLinks - 1).verts(i + n + 1).y
        c.Z = c.Z + st.Links(st.nLinks - 1).verts(i + n + 1).Z
    Next i
    c.x = c.x / (n + 1)
    c.y = c.y / (n + 1)
    c.Z = c.Z / (n + 1)
    m3StringHeadCenter = c
End Function

Public Function m3StringMidle(ByRef st As Strings) As Integer
    Dim c1 As m3Point
    Dim c2 As m3Point
    Dim P1 As m3Point
    Dim P2 As m3Point
    Dim P3 As m3Point
    Dim n As Vector
    Dim i As Integer
    Dim mid As Integer
    mid = st.nStrings - 1
    For i = 0 To st.nStrings - 2
        c1 = m3StringGuitarCenter(st.Strings(i))
        c2 = m3StringGuitarCenter(st.Strings(i + 1))
        n = m3ops.m3VectorInit(c1, c2)
        P1.x = (c1.x + c2.x) / 2
        P1.y = (c1.y + c2.y) / 2
        P1.Z = (c1.Z + c2.Z) / 2
        c1 = m3StringHeadCenter(st.Strings(i))
        c2 = m3StringHeadCenter(st.Strings(i + 1))
        P2.x = (c1.x + c2.x) / 2
        P2.y = (c1.y + c2.y) / 2
        P2.Z = (c1.Z + c2.Z) / 2
        n = m3ops.m3VectorCross(n, m3VectorInit(P1, P2))
        m3ops.m3VectorSetLength 10, n
        P3 = m3ops.m3PointAddVector(n, P2)
        If Not m3draw.m3PlaneIsVisible(P1, P2, P3) Then
            mid = i
            Exit For
        End If
        
    Next i
    m3StringMidle = mid
End Function

Public Function m3StringsInit(ByVal nStrings As Integer, ByVal Color As Long) As Strings
    Dim s As Strings
    Dim st As m3StringType
    Dim m As m3Matrix
    Dim i As Integer
    s.nStrings = nStrings
    ReDim s.Strings(s.nStrings - 1) As m3StringType
    st = m3StringInit(0.041 * R, R * 6, 32)
    For i = 0 To s.nStrings - 1
        m = m3ops.m3Translate(-8 * i, 0, 0)
        s.Strings(i) = m3StringInit(st.R - 0.5 * i, st.l, st.nLinks)
        s.Color = Color
        m3StringApply s.Strings(i), m
    Next i
    m3StringsInit = s
End Function

Public Sub m3StringsDraw(ByRef obj As Object, ByRef s As Strings)
    Dim i As Integer
    Dim j As Integer
    Dim zi As Double
    Dim zj As Double
    Dim temp As Double
    Dim mid As Integer
    Dim oldColor As Long
    oldColor = obj.FillColor
    mid = m3StringMidle(s)
    
    For i = 0 To mid - 1
        obj.FillColor = s.Color
        m3StringDraw obj, s.Strings(i)
    Next i
    For i = s.nStrings - 1 To mid Step -1
        obj.FillColor = s.Color
        m3StringDraw obj, s.Strings(i)
    Next i
    obj.FillColor = oldColor
End Sub

Public Sub m3StringsApply(ByRef s As Strings, ByRef m As m3Matrix)
    Dim i As Integer
    For i = 0 To s.nStrings - 1
        m3StringApply s.Strings(i), m
    Next i
End Sub

Public Function m3StringsCenter(ByRef s As Strings) As m3Point
    Dim cP As m3Point
    Dim Sp() As m3Point
    Dim i As Integer
    ReDim Sp(s.nStrings - 1) As m3Point
    cP.x = 0: cP.y = 0: cP.Z = 0
    For i = 0 To s.nStrings - 1
        Sp(i).x = 0: Sp(i).y = 0: Sp(i).Z = 0
    Next i
    For i = 0 To s.nStrings - 1
        Sp(i) = m3StringCenter(s.Strings(i))
    Next i
    For i = 0 To s.nStrings - 1
        cP.x = cP.x + Sp(i).x
        cP.y = cP.y + Sp(i).y
        cP.Z = cP.Z + Sp(i).Z
    Next i
    cP.x = cP.x / s.nStrings
    cP.y = cP.y / s.nStrings
    cP.Z = cP.Z / s.nStrings
    m3StringsCenter = cP
End Function

Private Sub DWalls(ByRef obj As Object, ByRef G As Guitar, ByRef s As Strings)
    Dim left As Boolean ' wall(0)
    Dim right As Boolean ' wall(1)
    Dim oldc As Long
    left = m3PlaneIsVisible(G.Wall(0).verts(1), G.Wall(0).verts(2), G.Wall(0).verts(3))
    right = m3PlaneIsVisible(G.Wall(1).verts(3), G.Wall(1).verts(2), G.Wall(1).verts(1))
    oldc = obj.FillColor
    If left And right Then
        obj.FillColor = G.cWallG
        m3Solid.SolidDrawShading obj, G.Wall(0)
        obj.FillColor = G.cWallH
        m3Solid.SolidDrawShading obj, G.Wall(1)
        m3StringsDraw obj, s
        Exit Sub
    End If
    If left Then
        obj.FillColor = G.cWallG
        m3Solid.SolidDrawShading obj, G.Wall(0)
        m3StringsDraw obj, s
        obj.FillColor = G.cWallH
        m3Solid.SolidDrawShading obj, G.Wall(1)
        Exit Sub
    End If
    If right Then
        obj.FillColor = G.cWallH
        m3Solid.SolidDrawShading obj, G.Wall(1)
        m3StringsDraw obj, s
        obj.FillColor = G.cWallG
        m3Solid.SolidDrawShading obj, G.Wall(0)
        Exit Sub
    End If
    m3StringsDraw obj, s
    obj.FillColor = G.cWallG
    m3Solid.SolidDrawShading obj, G.Wall(0)
    obj.FillColor = G.cWallH
    m3Solid.SolidDrawShading obj, G.Wall(1)
    obj.FillColor = oldc
End Sub

Private Sub DWallsWPickup(ByRef obj As Object, ByRef G As Guitar, ByRef s As Strings)
    If m3PlaneIsVisible(G.Wall(1).verts(0), G.Wall(1).verts(1), G.Wall(1).verts(5)) Then
        PickupDraw obj, G
        DWalls obj, G, s
    Else
        DWalls obj, G, s
        PickupDraw obj, G
    End If
End Sub

Private Function PickingCilinder(ByRef p0 As m3Point, ByRef v As Vector, ByRef P1 As m3Point, ByRef P2 As m3Point, ByVal R As Double) As Double
    'פעולה זו קובעת האם יש חיתוך בין הוקטור הנתון לגליל
    Dim a As Double
    Dim B As Double
    Dim x As Double
    Dim Z As Double
    Dim c As Double
    Dim y As Double
    Dim T() As Double
    ' Cylidrical surface
    Dim t1 As Double
    ' Up base
    Dim t2 As Double
    ' Down Base
    Dim t3 As Double
    Dim min As Double
    Dim m As m3Matrix
    Dim cilVector As Vector
    m3VectorSetLength 1, v
    ' setting points
    cilVector = m3VectorInit(P1, P2)
    m = m3ops.m3VectorToY(P1, cilVector)
    m3Pointapply p0, m
    m3Pointapply P1, m
    m3Pointapply P2, m
    m3VectorApply v, m
        
    ' calcs
    a = v.x * v.x + v.Z * v.Z
    B = 2 * (p0.x * v.x + p0.Z * v.Z)
    c = p0.x * p0.x + p0.Z * p0.Z - R * R
    T = sqrEquasion(a, B, c)
    t1 = NoCollision
    If UBound(T) > 0 Then
        If T(0) > 0 Then
            y = p0.y + v.y * T(0)
            If y <= P2.y And y >= 0 Then
                t1 = T(0)
            End If
        
        End If
    End If
    If Abs(v.y) < 0.0000001 Then
        t2 = NoCollision
        t3 = NoCollision
    Else
        t2 = (P2.y - p0.y) / v.y
        If t2 >= 0 Then
           x = p0.x + v.x * t2
           Z = p0.Z + v.Z * t2
           If x * x + Z * Z > R * R Then
               t2 = NoCollision
           End If
        Else
            t2 = NoCollision
        End If
    End If
    If t3 > NoCollision Then
       t3 = -p0.y / v.y
        If t3 >= 0 Then
           x = p0.x + v.x * t3
           Z = p0.Z + v.Z * t3
           If x * x + Z * Z > R * R Then
               t3 = NoCollision
           End If
        Else
            t3 = NoCollision
        End If
    End If
    min = NoCollision
    If t1 >= 0 Then
        min = t1
    End If
    If min = NoCollision Then
        If t2 >= 0 Then
            min = t2
        End If
    Else
        If t2 >= 0 And t2 < min Then
            min = t2
        End If
    End If
    
    If min = NoCollision Then
        If t3 >= 0 Then
            min = t3
        End If
    Else
        If t3 >= 0 And t3 < min Then
            min = t3
        End If
    End If
    PickingCilinder = min
End Function

Public Function StringPick(ByVal x As Double, ByVal y As Double, ByRef s As Strings) As Integer
    'פעולה זו קובעת נלחץ מיתר, ע"י חיתוך בין וקטור למיתר.
    Dim i As Integer
    Dim P1 As m3Point
    Dim P2 As m3Point
    Dim p As m3Point
    Dim StartP As m3Point
    Dim v As Vector
    Dim m As m3Matrix
    Dim R As Double
    Dim mouseV As Vector
    Dim dist As Double
    Dim T As Double
    dist = m3getdistance
    For i = 0 To 3
        StartP = m3PointInit(0, 0, dist)
        mouseV = m3VectorInit(StartP, m3PointInit(x, y, 0))
        P1 = m3Guitar.m3StringGuitarCenter(s.Strings(i))
        P2 = m3Guitar.m3StringHeadCenter(s.Strings(i))
        R = m3VectorLength(m3VectorInit(P1, s.Strings(i).Links(0).verts(0)))
        T = PickingCilinder(StartP, mouseV, P1, P2, R)
        If T <> NoCollision Then
            StringPick = i
            Exit Function
        End If
    Next i
    StringPick = -1
End Function

