VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Undo\Redo"
   ClientHeight    =   2535
   ClientLeft      =   180
   ClientTop       =   525
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmUndo.frx":0000
      Top             =   0
      Width           =   4335
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "Edit"
      Begin VB.Menu mnu_undo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_redo 
         Caption         =   "Redo"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By     Jim Jose
'email  jimjosev33@yahoo.com
Option Explicit    'declaring a new  Collection
Dim UndoStack As New Collection    'It determines the position
Dim Pos As Integer
Private Sub Form_Load()
    MsgBox "Err.Description, I Need your feed back", vbCritical
End Sub
Private Sub mnu_redo_Click()
    If Pos < UndoStack.Count Then Pos = Pos + 1     'avoidind 'index'  error
    Text1 = UndoStack(Pos)      'Getting the text from the stack
End Sub
Private Sub mnu_undo_Click()
    If Pos = UndoStack.Count And Not UndoStack(UndoStack.Count) = Text1 Then Text1_KeyPress 0: mnu_undo_Click      'Resetting the stack with the current text as the last index
    Text1.Text = UndoStack(Pos)     'Getting the text from the stack
    If Not Pos <= 1 Then Pos = Pos - 1       'resetting position to the next data
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If UndoStack.Count > 100 Then UndoStack.Remove (1)  'Removing  the first data when the 'allowed maximum' reaches
    UndoStack.Add Text1.Text    'Setting the text to the stack
    Pos = UndoStack.Count       'Updating position
End Sub


