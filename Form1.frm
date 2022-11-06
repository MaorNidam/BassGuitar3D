VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuParallel 
         Caption         =   "Parallel"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPerspective 
         Caption         =   "Perspective"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SceneInit()

End Sub
Private Sub SceneApply()

End Sub
Private Sub SceneDraw()

End Sub
Private Sub mnuParallel_Click()
    mnuParallel.Checked = True
    mnuPerspective.Checked = False
    'm3draw.
    
End Sub

Private Sub mnuPerspective_Click()
    mnuParallel.Checked = False
    mnuPerspective.Checked = True
End Sub
