VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10080
   FillColor       =   &H00004040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   672
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddf 
      Caption         =   "Add"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbString 
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Text            =   "String"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbFret 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Text            =   "Fret"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Piece"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.ComboBox cmbScale 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   2880
      List            =   "frmMain.frx":0002
      TabIndex        =   7
      Text            =   "Scale"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.ComboBox cmbNote 
      Height          =   315
      ItemData        =   "frmMain.frx":0004
      Left            =   1440
      List            =   "frmMain.frx":0006
      TabIndex        =   6
      Text            =   "Note"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   6600
      Width           =   735
   End
   Begin VB.PictureBox picClef 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   8640
      Picture         =   "frmMain.frx":0008
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picTav 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   1
      Left            =   8280
      Picture         =   "frmMain.frx":0476
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picTav 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   0
      Left            =   7920
      Picture         =   "frmMain.frx":0808
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picNotes 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   1
      Top             =   5160
      Width           =   10095
   End
   Begin VB.PictureBox picdraw 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   4935
      Left            =   0
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   0
      Top             =   0
      Width           =   10095
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuNotes 
         Caption         =   "Notes"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFret 
         Caption         =   "Fret"
      End
   End
   Begin VB.Menu mnuPieces 
      Caption         =   "Pieces"
      Begin VB.Menu mnuPre 
         Caption         =   "Bach - Prelude (G minor)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MSolid As m3Matrix
Private xstart As Single
Private ystart As Single
Private P1 As m3Point
Private P2 As m3Point
Private G As Guitar
Private st As Strings
Private s As Strings
Private Const maxPiece = 32
Private p As Piece
Private amp0 As Double
Private phaza As Integer
Private Const Range = 2#
Private Const Nphazas = 30
Private Const alpha = 2 * Pie / Nphazas
Private Semaphore As Boolean
Private Stp As Boolean
Private Const NoCollision = -100
Private mone As Integer
' הצהרה על הפונקציה ליצירת חלק שקוף בתמונה
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private Const Shesh = 0
Private Const Reva = 1

Private Const space = 10
Private Notes(7) As Integer
Private Const a = 0
Private Const B = 1
Private Const c = 2
Private Const d = 3
Private Const E = 4
Private Const f = 5
Private Const Gi = 6
Private StopSound As Boolean
Private Sub SceneInit()
    m3draw.m3SetPerspective
    picdraw.ScaleLeft = -picdraw.ScaleWidth / 2
    picdraw.ScaleHeight = -picdraw.ScaleHeight
    picdraw.ScaleTop = Abs(picdraw.ScaleHeight / 2)
    s = m3StringsInit(4, RGB(150, 150, 158))
    G = GuitarInit(s)
    InitLightVector
    MSolid = m3Identity
    m3StringsApply s, MSolid
    GuitarApply G, MSolid
    m3draw.m3SetPerspective
    phaza = 0
    amp0 = 0
    mone = 0
    ReDim p.Notes(maxPiece) As Tav
    Semaphore = False
    
    cmbScale.AddItem "1 / 16", Shesh
    cmbScale.AddItem "1 / 4", Reva

    
    cmbNote.AddItem "A", a
    cmbNote.AddItem "B", B
    cmbNote.AddItem "C", c
    cmbNote.AddItem "D", d
    cmbNote.AddItem "E", E
    cmbNote.AddItem "F", f
    cmbNote.AddItem "G", Gi
    
    cmbFret.AddItem "0", 0
    cmbFret.AddItem "1", 1
    cmbFret.AddItem "2", 2
    cmbFret.AddItem "3", 3
    cmbFret.AddItem "4", 4
    cmbFret.AddItem "5", 5
    cmbFret.AddItem "6", 6
    cmbFret.AddItem "7", 7
    cmbFret.AddItem "8", 8
    cmbFret.AddItem "9", 9
    cmbFret.AddItem "10", 10
    cmbFret.AddItem "11", 11
    cmbFret.AddItem "12", 12
    
    cmbString.AddItem "E", 0
    cmbString.AddItem "A", 1
    cmbString.AddItem "D", 2
    cmbString.AddItem "G", 3

    Notes(a) = picTav(0).Height
    Notes(B) = picTav(0).Height - space / 2
    Notes(c) = picTav(0).Height - 2 * space / 2
    Notes(d) = picTav(0).Height - 3 * space / 2
    Notes(E) = picTav(0).Height - 4 * space / 2
    Notes(f) = picTav(0).Height - 5 * space / 2
    Notes(Gi) = picTav(0).Height - 6 * space / 2
End Sub
Private Sub SceneApply()
    GuitarApply G, MSolid
    m3StringsApply s, MSolid
    m3StringsApply st, MSolid
    
    If Not m3PlaneIsVisible(G.Pickup(2).verts(0), G.Pickup(2).verts(1), G.Pickup(2).verts(5)) Then
        Semaphore = True
    Else
        Semaphore = False
    End If
End Sub
Private Sub SceneDraw()
    picdraw.Cls
    GuitarDraw picdraw, G, s
    picdraw.Refresh
    
    Dim i As Integer
    For i = 0 To 4
        picNotes.Line (0, picTav(0).Height - space + i * space)-(picNotes.ScaleWidth, picTav(0).Height - space + i * space)
    Next i
    TransparentBlt picNotes.hdc, 0, picTav(0).Height - 14, picClef.ScaleWidth, picClef.ScaleHeight, picClef.hdc, 0, 0, picClef.ScaleWidth, picClef.ScaleHeight, vbWhite
    
    picNotes.Refresh
End Sub
Private Sub StringStrum(ByVal nString As Integer, ByVal fret As Integer, ByVal isShesh As Boolean)
    'פונקציה המקבלת ערכים לגבי פריטה על המיתר ומציירת פריטה על ידי תנועה
    'הרמונית של החלקים הגליים במיתר.
    Dim amp As Double
    Dim j As Integer
    Dim i As Integer
    Semaphore = True
    st = s
    StopSound = False
    i = (s.Strings(nString).nLinks - fret) / 2
    On Error GoTo errHandler
    If isShesh Then
        MediaPlayer1.FileName = App.Path & "\sounds\a" & nString & i & ".mp3"
    Else
        MediaPlayer1.FileName = App.Path & "\sounds\" & nString & i & ".mp3"
    End If
    MediaPlayer1.Play
errHandler:
    For i = 0 To Nphazas * 50
        amp = Range * Sin(alpha * phaza)
        phaza = (phaza + 1) Mod Nphazas
        For j = 0 To fret / 2 - 1
            m3StringSwing s.Strings(nString), (amp - amp0) * 0.6 * (j / 10), j
            m3StringSwing s.Strings(nString), (amp - amp0) * 0.6 * (j / 10), fret - j - 1
            DoEvents
        Next j
        m3StringSwing s.Strings(nString), (amp - amp0) * 0.06, s.Strings(0).nLinks / 2
        amp0 = amp
        If StopSound Then Exit For
        SceneDraw
    Next i
    s = st
    SceneDraw
    phaza = 0
    amp0 = 0
    Semaphore = False
End Sub

Private Sub cmdAdd_Click()
    'לחיצה על לחצן זה מוסיפה תיו למחברת תוים, מנגנת את אותו התיו,
    'ומכניסה את התיו למערך.
    'פעולה זו עובדת לפי תווים.
    If Semaphore Then Exit Sub
    Dim i As Integer
    Dim j As Integer
    i = cmbScale.ListIndex
    j = cmbNote.ListIndex
    If i < 0 Or j = -1 Then
        MsgBox ("Invaild add.")
    Else
        If mone >= maxPiece Then
            MsgBox "Can't add more then 32 notes."
        Else
        cmdPlay.Enabled = True
        TransparentBlt picNotes.hdc, picClef.Width + picTav(i).ScaleWidth * mone, Notes(j), picTav(i).ScaleWidth, picTav(i).ScaleHeight, picTav(i).hdc, 0, 0, picTav(i).ScaleWidth, picTav(i).ScaleHeight, vbWhite
        mone = mone + 1
        Select Case j
            Case a
                p.nNotes = mone
                p.Notes(mone - 1).nString = 1
                p.Notes(mone - 1).fret = s.Strings(1).nLinks
                If i = 1 Then
                    StringStrum 1, s.Strings(1).nLinks, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 1, s.Strings(1).nLinks, True
                    p.Notes(mone - 1).isShesh = True
                End If
            Case B
                p.nNotes = mone
                p.Notes(mone - 1).nString = 1
                p.Notes(mone - 1).fret = s.Strings(1).nLinks - 4
                If i = 1 Then
                    StringStrum 1, s.Strings(1).nLinks - 4, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 1, s.Strings(1).nLinks - 4, True
                    p.Notes(mone - 1).isShesh = True
                End If
            Case c
                p.nNotes = mone
                p.Notes(mone - 1).nString = 1
                p.Notes(mone - 1).fret = s.Strings(1).nLinks - 6
                If i = 1 Then
                    StringStrum 1, s.Strings(1).nLinks - 6, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 1, s.Strings(1).nLinks - 6, True
                    p.Notes(mone - 1).isShesh = True
                End If
            Case d
                p.nNotes = mone
                p.Notes(mone - 1).nString = 2
                p.Notes(mone - 1).fret = s.Strings(2).nLinks
                If i = 1 Then
                    StringStrum 2, s.Strings(2).nLinks, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 2, s.Strings(2).nLinks, True
                    p.Notes(mone - 1).isShesh = True
                End If
            Case E
                p.nNotes = mone
                p.Notes(mone - 1).nString = 2
                p.Notes(mone - 1).fret = s.Strings(2).nLinks - 4
                If i = 1 Then
                    StringStrum 2, s.Strings(2).nLinks - 4, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 2, s.Strings(2).nLinks - 4, True
                    p.Notes(mone - 1).isShesh = True
                End If
            Case f
                p.nNotes = mone
                p.Notes(mone - 1).nString = 2
                p.Notes(mone - 1).fret = s.Strings(2).nLinks - 6
                If i = 1 Then
                    StringStrum 2, s.Strings(2).nLinks - 6, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 2, s.Strings(2).nLinks - 6, True
                    p.Notes(mone - 1).isShesh = True
                End If
            Case Gi
                p.nNotes = mone
                p.Notes(mone - 1).nString = 3
                p.Notes(mone - 1).fret = s.Strings(3).nLinks
                If i = 1 Then
                    StringStrum 3, s.Strings(3).nLinks, False
                    p.Notes(mone - 1).isShesh = False
                Else
                    StringStrum 3, s.Strings(3).nLinks, True
                    p.Notes(mone - 1).isShesh = True
                End If
        End Select
        End If
    End If
    picNotes.Refresh
End Sub

Private Sub cmdAddf_Click()
    'לחיצה על לחצן זה מנגנת את התיו הנבחר על ידי הפוזיציה במיתר הנבחר,
    'ומכניסה את התיו למערך. (פעולה זו אינה מציירת תיו במחברת)
    'פעולה זו עובדת לפי פוזיציה על המיתר.
    If Semaphore Then Exit Sub
    Dim i As Integer
    Dim j As Integer
    i = cmbFret.ListIndex
    j = cmbString.ListIndex
    If i < 0 Or j = -1 Then
        MsgBox ("Invaild add.")
    Else
        If mone >= maxPiece Then
            MsgBox "Can't add more then 32 notes."
        Else
        cmdPlay.Enabled = True
        mone = mone + 1
        p.nNotes = mone
        p.Notes(mone - 1).fret = s.Strings(j).nLinks - i * 2
        p.Notes(mone - 1).nString = j
        If cmbScale.ListIndex = 0 Then
            StringStrum j, s.Strings(j).nLinks - i * 2, True
            p.Notes(mone - 1).isShesh = True
        Else
            StringStrum j, s.Strings(j).nLinks - i * 2, False
            p.Notes(mone - 1).isShesh = False
        End If
        End If
    End If
    picNotes.Refresh
End Sub

Private Sub cmdPlay_Click()
    'פעולה זו מנגנת את התווים שנבחרו בזה אחר זה.
    Dim i As Integer
    cmdStop.Enabled = True
    For i = 0 To p.nNotes - 1
        If Stp = True Then
            Exit Sub
        End If
        StringStrum p.Notes(i).nString, p.Notes(i).fret, p.Notes(i).isShesh
    Next i
    cmdStop.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Stp = True
End Sub

Private Sub Form_Load()
    Me.Caption = "Bass Player"
    Me.Show
    SceneInit
    SceneDraw
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
    StopSound = True
End Sub

Private Sub mnuFret_Click()
    mnuFret.Checked = True
    mnuNotes.Checked = False
    cmbFret.Visible = True
    cmbNote.Visible = False
    cmbString.Visible = True
    cmdAddf.Visible = True
    cmdAdd.Visible = False
    ReDim p.Notes(maxPiece) As Tav
    p.nNotes = 0
    MsgBox "Adding by fret won't add notes to the notepad." & vbCrLf & "Please note that your piece has been restarted."
End Sub

Private Sub mnuNotes_Click()
    mnuFret.Checked = False
    mnuNotes.Checked = True
    cmbFret.Visible = False
    cmbNote.Visible = True
    cmbString.Visible = False
    cmdAddf.Visible = False
    cmdAdd.Visible = True
    ReDim p.Notes(maxPiece) As Tav
    p.nNotes = 0
    MsgBox "Please note that your piece has been restarted."
End Sub

Private Sub mnuPre_Click()
    'פעולה זו מנגנת את הפרילוד מהסוויטה הראשונה של באך.
    Dim tmp As Piece
    Dim i As Integer
    tmp = p
    cmdStop.Enabled = True
    p = PieceInit(App.Path & "\Piece\Prelude.dat")
    For i = 0 To p.nNotes - 1
        If Stp = True Then
            p = tmp
            Exit Sub
        End If
        StringStrum p.Notes(i).nString, s.Strings(p.Notes(i).nString).nLinks - p.Notes(i).fret * 2, True
    Next i
    p = tmp
    cmdStop.Enabled = False
End Sub

Private Sub picdraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            MSolid = m3ops.m3Translate(-5, 0, 0)
        Case vbKeyRight
            MSolid = m3ops.m3Translate(5, 0, 0)
        Case vbKeyUp
            MSolid = m3ops.m3Translate(0, 5, 0)
        Case vbKeyDown
            MSolid = m3ops.m3Translate(0, -5, 0)
        Case vbKeyPageUp
            MSolid = m3ops.m3Translate(0, 0, 5)
        Case vbKeyPageDown
            MSolid = m3ops.m3Translate(0, 0, -5)
        Case Else
            Exit Sub
    End Select
    SceneApply
    SceneDraw
End Sub

Private Sub picdraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    xstart = x
    ystart = y
    If Semaphore Then
        Exit Sub
    End If
    If Button = vbRightButton Then
        If StringPick(x, y, s) <> -1 Then
            StringStrum StringPick(x, y, s), s.Strings(StringPick(x, y, s)).nLinks, False
        End If
        
    End If
End Sub

Private Sub picdraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dx As Single
    Dim dy As Single
    Dim cent As m3Point
    dx = x - xstart
    dy = y - ystart
    xstart = x
    ystart = y
    If Button = vbLeftButton Then
        cent = m3SolidCenter(G.Pickup(2))
        MSolid = m3ops.m3Translate(-cent.x, -cent.y, -cent.Z)
        MSolid = m3ops.m3MatMultiply(MSolid, m3RotateY(dx / 100))
        MSolid = m3ops.m3MatMultiply(MSolid, m3RotateX(-dy / 100))
        MSolid = m3ops.m3MatMultiply(MSolid, m3ops.m3Translate(cent.x, cent.y, cent.Z))
        SceneApply
        SceneDraw
    End If
End Sub

