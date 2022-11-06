Attribute VB_Name = "m3Sound"
Option Explicit
Public Type Tav
    nString As Integer
    fret As Integer
    isShesh As Boolean
End Type
Public Type Piece
    nNotes As Integer
    Notes() As Tav
End Type

Public Function PieceInit(ByVal Path As String) As Piece
    Dim p As Piece
    Dim FileNum As Single
    Dim i As Single
    FileNum = FreeFile
    Open Path For Input As #FileNum
    Input #FileNum, p.nNotes
    ReDim p.Notes(p.nNotes - 1) As Tav
    For i = 0 To p.nNotes - 1
        Input #FileNum, p.Notes(i).nString, p.Notes(i).fret, p.Notes(i).isShesh
    Next i
    Close #FileNum
    PieceInit = p
End Function
