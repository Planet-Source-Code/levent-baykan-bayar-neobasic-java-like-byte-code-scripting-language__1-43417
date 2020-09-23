VERSION 5.00
Begin VB.Form DebugW 
   Caption         =   "Debug Window [Press F1 to save,ESC to exit]"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3690
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
   End
End
Attribute VB_Name = "DebugW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
If KeyCode = vbKeyF1 Then
Open App.Path + "\debug" & Format(Time, "hh-mm-ss") & ".txt" For Output As #1
For i = 0 To List1.ListCount - 1
Print #1, List1.List(i)
Next
Close #1
MsgBox "Saved to " & vbCr & App.Path + "\debug" & Time & ".txt"

End If


End Sub

Private Sub Form_Resize()
List1.Width = Width - 200
List1.Height = Height - 500

End Sub
