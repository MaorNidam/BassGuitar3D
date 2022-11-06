Attribute VB_Name = "m3Polly"
Option Explicit
Public Type polly
    nVerts As Single
    verts() As m3Point
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
Public Const MaxPoints = 10000
Private pt(MaxPoints) As POINTAPI
Private xOrg As Single
Private yOrg As Single
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Function PollyInit(ByVal Path As String) As polly
    Dim p As polly
    Dim FileNum As Single
    Dim i As Single
    FileNum = FreeFile
    Open Path For Input As #FileNum
    Input #FileNum, p.nVerts
    ReDim p.verts(p.nVerts - 1) As m3Point
    For i = 0 To p.nVerts - 1
        Input #FileNum, p.verts(i).x, p.verts(i).y, p.verts(i).Z
    Next i
    Close #FileNum
    PollyInit = p
End Function

Public Sub m3polyApply(ByRef p As polly, ByRef m As m3Matrix)
    Dim i As Single
    For i = 0 To p.nVerts - 1
        m3Pointapply p.verts(i), m
    Next i
End Sub

Public Sub m3PolyDraw(ByRef obj As Object, ByRef p As polly, ByVal closed As Boolean)
    Dim i As Single
    For i = 0 To p.nVerts - 2
        m3linedraw obj, p.verts(i), p.verts(i + 1)
    Next i
    If (closed) Then
        m3linedraw obj, p.verts(p.nVerts - 1), p.verts(0)
    End If
End Sub

Public Function m3PolyCenter(ByRef p As polly) As m3Point
    Dim i As Single
    Dim cP As m3Point
    cP.x = 0: cP.y = 0: cP.Z = 0
    For i = 0 To p.nVerts - 1
        cP.x = cP.x + p.verts(i).x
        cP.y = cP.y + p.verts(i).y
        cP.Z = cP.Z + p.verts(i).Z
    Next i
    cP.x = cP.x / p.nVerts
    cP.y = cP.y / p.nVerts
    cP.Z = cP.Z / p.nVerts
    m3PolyCenter = cP
End Function

Public Sub m3PolyFill(ByRef obj As Object, ByRef p As polly)
    Dim f As Single
    Dim i As Single
    Dim dist As Single
    xOrg = -obj.ScaleLeft
    yOrg = obj.ScaleTop
    
    If m3draw.m3isparallel Then
        For i = 0 To p.nVerts - 1
            pt(i).x = xOrg + p.verts(i).x
            pt(i).y = yOrg - p.verts(i).y
        Next i
    Else
        dist = m3draw.m3getdistance
        For i = 0 To p.nVerts - 1
            
            f = dist / (dist - p.verts(i).Z)
            pt(i).x = xOrg + f * p.verts(i).x
            pt(i).y = yOrg - f * p.verts(i).y
        Next i
    End If
    
    Polygon obj.hdc, pt(0), p.nVerts
End Sub
