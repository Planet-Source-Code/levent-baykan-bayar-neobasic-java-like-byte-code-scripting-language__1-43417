VERSION 5.00
Begin VB.Form CForm 
   Caption         =   "Script Compiler and Executer"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Compiler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Tag             =   "no"
   Begin VB.CheckBox Check1 
      Caption         =   "Optimize"
      Height          =   195
      Left            =   4005
      TabIndex        =   7
      Top             =   45
      Width           =   2445
   End
   Begin VB.CommandButton Command8 
      Caption         =   "T"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2205
      Picture         =   "Compiler.frx":08CA
      TabIndex        =   6
      Top             =   45
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1755
      MaskColor       =   &H00008080&
      Picture         =   "Compiler.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1350
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Compiler.frx":0B0E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   450
      Picture         =   "Compiler.frx":0CF0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      Picture         =   "Compiler.frx":0DF2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   450
      Width           =   6405
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   4500
      Width           =   6405
   End
End
Attribute VB_Name = "CForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inS As String
Dim FName As String



Private Sub Command4_Click()
FName = CD(hWnd, "Scripts(*.txt)" + Chr(0) + "*.txt", "Open File", "*.txt")
If FName = "" Then Exit Sub
Open FName For Input As #1
inS = Input(LOF(1), 1)
Close #1
Text1 = inS

End Sub

Private Sub Command5_Click()
FName = CD(hWnd, "Scripts(*.txt)" + Chr(0) + "*.txt", "Save File", "*.txt")
If FName = "" Then Exit Sub
Open FName For Output As #1
Print #1, Text1.Text
Close #1

End Sub

Private Sub Command6_Click()
If Text1 = "" Then
MsgBox "Open a text file first", 16
Exit Sub
End If

'autosave
If FName <> "" Then
Open FName For Output As #1
Print #1, Text1
Close #1
End If

If Dir(FName + ".d") <> "" Then Kill FName + ".d"
CompilePiece Text1, "main"

End Sub

Private Sub Command7_Click()
Command6_Click

If Dir(FName + ".d") = "" Then

MsgBox "Open a bytecode file first" + vbCr + "then compile it", 16
Exit Sub




End If


Dim Run As New NeoBasicCompile
Run.ScriptFile = FName + ".d"
Set Run.ClientWindow = Form1
Run.RunBlock
End Sub

Private Sub Command8_Click()
If inS = "" Then
MsgBox "Open a text file first", 16
Exit Sub
End If

'Dim Run As New NeoBasic
'Run.Script = inS
'Set Run.ClientWindow = Form1
'Run.RunBlock
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Public Function CompilePiece(i As String, piece As String, Optional Parameter As String, Optional MyIndex As Integer, Optional TFormStr As String) As Boolean
    Dim j As Integer
    Dim kj As Integer
    Dim Msl As New NeoBasicCompile
    
    kj = FreeFile
    piece = LCase(piece)
exec:
    mn = InStr(1, i, piece + "(" + Parameter + ")")
    
    If mn = 0 Then RunPiece = False: Exit Function
    p1 = InStr(mn, i, "{")
    p2 = InStr(mn, i, "}")
    If p1 <> 0 And p2 = 0 Then MsgBox "Error in line!" + vbCr + "Missing { or } ", 16: Exit Function
    sctxt = Mid(i, p1 + 1, p2 - p1 - 1)
    sctxt = Replace(sctxt, vbCrLf, "")
    CompilePiece = True
    Open FName + ".d" For Binary As #1
    'Put #1, , Piece
    
    
    Msl.Script = sctxt
    
    Msl.DestFile = FName + ".d"
    
  
    Msl.Procedure = piece + "(" + Parameter + ")"
    
    Msl.Compile
    Close #1
Set Msl = Nothing
End Function
