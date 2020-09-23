VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function DllRegisterServer Lib "neodll.dll" () As Long
'Private Declare Function DllUnregisterServer Lib "neodll.dll" () As Long

Dim k As String
Dim t As String

Public Inclist As New Collection
Private GName As New Collection
Private GValue As New Collection

Private AName As New Collection
Private AValue As New Collection
Const EXE = 147456

Public Sub AddArray(ArName As String)
AName.Add ArName
AValue.Add ""

End Sub
Public Sub ChangeArr(ArName As String, Index As Integer, newV As Variant)
Dim oldArr As String
Dim tm() As String
Dim j As Integer

Dim i As Integer
For i = AName.Count To 1 Step -1
If LCase(AName.Item(i)) = LCase(ArName) Then
    AName.Remove i
    oldArr = AValue.Item(i)
    AValue.Remove i
    
    tm = Split(oldArr, "|")
    If Index > UBound(tm) Then ReDim Preserve tm(Index)
    tm(Index) = newV
    
        
    oldArr = ""
    AName.Add ArName
    For j = 0 To UBound(tm)
    oldArr = oldArr + tm(j) + "|"
    Next
    
    AValue.Add oldArr
    
    Exit Sub
End If
Next

End Sub
Public Function GetArr(ArName, Index As Integer)
Dim tm() As String
Dim j As Integer

Dim i As Integer

For i = 1 To AName.Count


If LCase(AName.Item(i)) = LCase(ArName) Then
    
    tm = Split(AValue.Item(i), "|")
    'MsgBox AValue.Item(i) & "   " & Index & "  " & tm(Index)
    If Index > UBound(tm) Then MsgBox "Invalid array index!" + vbCr + "Array : " & ArName & " Index: " & Index, 16: Exit Function
    GetArr = tm(Index) 'AValue.Item(i)
    
    Exit Function
End If
Next
GetArr = "45lbb"

End Function
Public Sub AddGlob(GlName As String)
GName.Add GlName
GValue.Add ""

End Sub
Public Sub ChangeGlob(GlName As String, newV As Variant)

Dim i As Integer
For i = GName.Count To 1 Step -1
If LCase(GName.Item(i)) = LCase(GlName) Then
    GName.Remove i
    GValue.Remove i
    
    
    GName.Add GlName
    GValue.Add newV
    Exit Sub
End If
Next

End Sub
Public Function GetGlob(GlName)

Dim i As Integer

For i = 1 To GName.Count


If LCase(GName.Item(i)) = LCase(GlName) Then
    GetGlob = GValue.Item(i)
    
    Exit Function
End If
Next
GetGlob = "45lbb"

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

RunPiece k, "onformkeydown", Trim(Str(KeyCode))
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RunPiece k, "onmouseup"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
'DllUnregisterServer

Set GValue = Nothing
Set GName = Nothing
Set AValue = Nothing
Set AName = Nothing
End
End Sub



Sub RunPiece(i As String, piece As String, Optional MyIndex As Variant)
Dim Run As New NeoBasicCompile
Run.Script = i
Set Run.ClientWindow = Form1
Run.RunBlock
End Sub
