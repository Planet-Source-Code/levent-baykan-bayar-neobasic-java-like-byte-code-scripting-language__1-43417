VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Sample ActiveX structure for NeoBasic
'Try adding functions,but do not forget
'to add their names to Funcs Sub.


Public ClientWindow As Form
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Function ExecuteFunction(NameOfFunction As String, arg() As String) As Variant
Dim TFN As Long
Dim buf As String
Dim arf As String
Dim f As Integer

On Error GoTo ErrorEx:




Select Case LCase(NameOfFunction)
    Case "fopen"
    TFN = FreeFile
        Select Case arg(1)
        Case "r"
        Open arg(0) For Binary As #TFN
        Case "w"
        Open arg(0) For Binary As #TFN
        End Select
        
        ExecuteFunction = TFN
    Case "fread"
        buf = String(Val(arg(1)), Chr(0))
        Get #Val(arg(0)), , buf
        ExecuteFunction = buf
        
    Case "fwrite"
        Put #Val(arg(0)), , arg(1)
    Case "lof"
        ExecuteFunction = LOF(Val(arg(0)))
    Case "eof"
        ExecuteFunction = EOF(Val(arg(0)))
    Case "seek"
        If UBound(arg) = 1 Then
        Seek #Val(arg(0)), arg(1)
        Else
        ExecuteFunction = Seek(arg(0))
        End If
    Case "filelen"
        ExecuteFunction = FileLen(arg(0))
    Case "fclose"
        Close #Val(arg(0))
    Case "fdelete"
        Kill arg(0)
    Case "freefile"
        ExecuteFunction = FreeFile
End Select
Exit Function
ErrorEx:
 On Error Resume Next
 For f = 0 To UBound(arg)
 arf = arf + arg(f) + ","
 Next
 arf = "(" + Left(arf, Len(arf) - 1) + ")"
 MsgBox "Error Module [io.dll]" + vbCr + "Error Procedure [" + NameOfFunction + arf + "]" + vbCr + "VB Error :" + Err.Description, 16
 ExecuteFunction = "Error" & Err.Number
 
End Function

Public Function Funcs() As String
Funcs = "fopen;fread;fwrite;lof;eof;seek;filelen;fclose;fdelete;freefile"
End Function
