Attribute VB_Name = "m3Solid"
Option Explicit
Public Type Face3D
    nVerts As Integer
    Face() As Integer
End Type
Public Type Soild
    nVerts As Integer
    nFaces As Integer
    Verts() As m3Point
    Faces() As Face3D
End Type

Public Function m3SolidInit(ByVal Path As String) As Soild
    Dim s As Soild
    Dim FileNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    FileNum = FreeFile
    Open Path For Input As #FileNum
    Input #FileNum, s.nVerts
    Input #FileNum, s.nFaces
    ReDim s.Verts(s.nVerts - 1) As m3Point
    For i = 0 To s.nVerts - 1
        Input #FileNum, s.Verts(i).x, s.Verts(i).y, s.Verts(i).Z
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

Public Sub m3SolidApply(ByRef s As Soild, ByRef m As m3Matrix)
    Dim i As Integer
    For i = 0 To s.nVerts - 1
        m3Pointapply s.Verts(i), m
    Next i
End Sub

Public Sub SoildDraw(ByRef obj As Object, ByRef s As Soild)
    
End Sub

Public Function m3SolidToString(ByRef s As Soild) As String
    Dim x As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    x = s.nVerts & vbTab & s.nFaces & vbCrLf
    For i = 0 To s.nVerts - 1
        x = x & m3ops.m3PointToString(s.Verts(i)) & vbCrLf
    Next i
    For j = 0 To s.nFaces - 1
        x = x & s.Faces(j).nVerts
        For k = 0 To s.Faces(j).nVerts - 1
            x = x & vbTab & s.Faces(j).Face(k)
        Next k
        x = x & vbCrLf
    Next j
    m3SolidToString = x
End Function
