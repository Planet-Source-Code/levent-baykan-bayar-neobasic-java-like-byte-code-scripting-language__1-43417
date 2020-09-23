VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NeoBasicCompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Var
    VType As String * 3
    Name As String
    value As Variant
End Type


Public Script As String ' holds script text

Private BreakScript As Boolean 'break running

Public ScriptFile As String 'file name of script
Public DestFile As String 'file name of byte code
Public Procedure As String 'name of procedure main()..etc

Public ClientWindow As Form '

Public MouseX As Integer 'unused in this
Public MouseY As Integer 'unused in this
Public Mbutton As Byte 'unused in this
Public DMode As Boolean 'switchs debug mode
Public CmdL As String 'Command$ passer


Dim Incs(15) As Object 'ActiveX Include list


Dim SkipQU As Byte 'evaluation sub Quote holder

Dim Vars(63) As Var 'local vars
Dim Ident(45) As String 'names of ActiveX DLLs





Dim Parse() As Segment 'Byte-Code data holder
Dim Parser As Segment 'Byte-Code data holder for compiler

Dim ECount As Integer 'Error Count
Dim WCount As Integer 'Warning count ,not used

Dim n As Integer 'Line number


Dim NumCurrentLine As Integer 'Line number
Private Sub Init()
    'Initialization of engine
    
    
    MyF 'loads name of built in functions see global.bas
    
    If Procedure = "main()" Then Exit Sub
    DebugAdd "$INIT"
    Dim OldUB As Integer
    Dim j As Integer
    Dim f As Integer
    Dim tmpF() As String
    'Below is for other subs in script such as onLoad()
    'on Exit()...but not yet implemented
    
    'For f = 1 To ClientWindow.Inclist.Count
    
    'Set Incs(f - 1) = CreateObject(ClientWindow.Inclist.Item(f) + ".Func")
    'Set Incs(f - 1).ClientWindow = ClientWindow
    'OldUB = UBound(MyFuncs)
    'tmpF = Split(Incs(f - 1).Funcs, ";")
    'ReDim Preserve MyFuncs(OldUB + UBound(tmpF) + 1)
    '
    '                For j = 0 To UBound(tmpF)
    '                MyFuncs(OldUB + j + 1) = tmpF(j)
    '                Next
    'Next
End Sub


Public Sub RunBlock()
    '==========================================================
    'Yeah,This is the main sub which runs every segment of
    'byte code this is basically a for/next loop,for better
    'management.
    '==========================================================
    
    
    '=========================
    
    Init
    'On Error GoTo erc
   
    Dim lCrLfSpot As Long, lCrLfSpotOld As Long
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer, t As Integer
    Dim sTemp As String, sTemp2 As String
    Dim arTemp() As String, arTemp2() As String
    Dim tmpGlob As String
    Dim GoToLabel As String
    Dim Skip As Boolean
    Dim IfSkip As Boolean
    Dim ForSkip As Boolean
    
    
    Dim Fok As Byte 'fonksiyon
    
    Dim tmpF() As String
    Dim OldUB As Integer
    
    
    Dim p As Byte
    Dim tmpEq As String
    Dim tmpAss As String
    Dim j As Integer
    Dim iq As Boolean
    Dim TmpChar As String * 1
    Dim o As Byte
    Dim CurVar As Byte
    Dim st
    Dim et
    Dim forT As String, forTo As String, forTo2 As String
    
    Dim forI As Integer, forIn As Integer
    Dim Nline As Integer
    Dim olP As Long
    'Open File----------------
    Open ScriptFile For Binary As #1
    Get #1, , Nline
    ReDim Parse(Nline)
    Get #1, , Parse
    'MsgBox "File cmp " & Seek(1) & " " & LOF(1)
    Close #1
    
    'Debug the byte-code into run.txt
    
    Open App.Path + "\run.txt" For Output As #2
    Dim WhileLoop As New Collection
    Dim IfWhileLoop As New Collection
    
    For NumCurrentLine = 0 To Nline
        If BreakScript Then Exit For
        'olP = Seek(1)
        'Get #1, , Parse(NumCurrentLine)
        If Parse(NumCurrentLine).Comm = 0 Then GoTo SkipThisSegment:
        'MsgBox NumCurrentLine & " [" & Parse(NumCurrentLine).Comm & ":" & MyData(Parse(NumCurrentLine).Comm - 1) & "] => " & DecStr
        
        
        Print #2, MyData(Parse(NumCurrentLine).Comm - 1) & " " & DecStr
        
        
        If GoToLabel <> "" Then 'Check Label
            If Parse(NumCurrentLine).Comm = 1 Then
                GoToLabel = ""
            End If
            GoTo SkipThisSegment:
        End If
        
        
        If Skip = True Then 'Check end of while block
            If Parse(NumCurrentLine).Comm = 2 Then
                Skip = False
            End If
            GoTo SkipThisSegment:
        End If
        '
        If ForSkip = True Then 'Check end of for block
            If Parse(NumCurrentLine).Comm = 3 Then
                ForSkip = False
            End If
            GoTo SkipThisSegment:
        End If
        
        
        If IfSkip = True Then ' Check end of if block
            
            Select Case Parse(NumCurrentLine).Comm
            Case 4
                IfSkip = False
            Case 5
                IfSkip = False
            End Select
            GoTo SkipThisSegment:
        Else
            If Parse(NumCurrentLine).Comm = 5 Then
                Skip = True
                GoTo SkipThisSegment:
            End If
        End If
        
        Select Case Parse(NumCurrentLine).Comm
            
            
        Case 2
            If WhileLoop.Count > 0 Then
                
                'Debug.Print Mid(WhileLoop(WhileLoop.Count), InStr(1, WhileLoop(WhileLoop.Count), ":") + 1), 0, "wend"
                If Eval(Mid(WhileLoop(WhileLoop.Count), InStr(1, WhileLoop(WhileLoop.Count), ":") + 1)) Then
                    NumCurrentLine = CInt(Mid(WhileLoop(WhileLoop.Count), 1, InStr(1, WhileLoop(WhileLoop.Count), ":") - 1)) - 1
                    'MsgBox "NumCurrentLine :" & NumCurrentLine & vbCr & DecStr, 0, "Wend"
                    'Seek #1, olP
                Else
                    WhileLoop.Remove WhileLoop.Count
                End If
                GoTo SkipThisSegment:
            End If
            
            
        Case 4
            If IfWhileLoop.Count > 0 Then
                'Looping for if block
                
                If Eval(Mid(IfWhileLoop(IfWhileLoop.Count), InStr(1, IfWhileLoop(IfWhileLoop.Count), ":") + 1)) Then
                    NumCurrentLine = CInt(Mid(IfWhileLoop(IfWhileLoop.Count), 1, InStr(1, IfWhileLoop(IfWhileLoop.Count), ":") - 1)) - 1
                    
                Else
                    IfWhileLoop.Remove IfWhileLoop.Count
                End If
                
            End If
            
            GoTo SkipThisSegment:
            
            
        Case 9 'if block
            
            If Eval(DecStr) Then
                IfWhileLoop.Add NumCurrentLine + 1 & ":" & Trim(DecStr)
                
            Else
                IfSkip = True
                
            End If
            
            
            GoTo SkipThisSegment:
            
        Case 10 ' Goto a label
            GoToLabel = DecStr
            
            GoTo SkipThisSegment:
        Case 11 'Declare a global variable
            
            tmpGlob = DecStr
            
            ClientWindow.AddGlob Trim(tmpGlob)
            
            GoTo SkipThisSegment:
        Case 12
            tmpGlob = DecStr
            ClientWindow.AddArray Trim(tmpGlob)
            GoTo SkipThisSegment:
            
        Case 13 'Include directive
            tmpGlob = DecStr
            If Dir(DecStr + ".dll") = "" Then
                MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Lib can't be found! ", 16
                Exit Sub
            End If
            
            If DecStr = "" Then
                MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "No lib specicifed with include! ", 16
                Exit Sub
            End If
            
            If ClientWindow.Inclist.Count = 17 Then
                MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "You can't include more than 16 modules in same scope!", 16
                Exit Sub
            End If
            
            For j = 1 To ClientWindow.Inclist.Count
                If ClientWindow.Inclist(j) = tmpGlob Then GoTo SkipThisSegment:
            Next
            
            ClientWindow.Inclist.Add tmpGlob
            'register DLLs
            Shell "regsvr32.exe /s " & tmpGlob + ".dll", vbHide
            Set Incs(ClientWindow.Inclist.Count - 1) = CreateObject(ClientWindow.Inclist.Item(ClientWindow.Inclist.Count) + ".Func")
            Set Incs(ClientWindow.Inclist.Count - 1).ClientWindow = ClientWindow
            
            OldUB = UBound(MyFuncs)
            'Add function names in dll to our list
            tmpF = Split(Incs(ClientWindow.Inclist.Count - 1).Funcs, ";")
            
            ReDim Preserve MyFuncs(OldUB + UBound(tmpF) + 1)
            
            For j = 0 To UBound(tmpF)
                MyFuncs(OldUB + j + 1) = tmpF(j)
            Next
            
            
            
            
        Case 14 ' Declare an integer variable
            
            tmpGlob = DecStr
            'If Mid(tmpGlob, 1, 1) <> "%" Then MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Variable Declaration error!" + vbCr + "You must start a variable name with %", 16
            Vars(CurVar).Name = tmpGlob
            Vars(CurVar).value = 0
            Vars(CurVar).VType = "int"
            CurVar = CurVar + 1
            GoTo SkipThisSegment:
        Case 15 ' Declare a string variable
            tmpGlob = DecStr
            'If Mid(tmpGlob, 1, 1) <> "%" Then MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Variable Declaration error!" + vbCr + "You must start a variable name with %", 16
            Vars(CurVar).Name = tmpGlob
            Vars(CurVar).value = CStr(Vars(CurVar).value)
            Vars(CurVar).VType = "str"
            CurVar = CurVar + 1
            
            GoTo SkipThisSegment:
        Case 16 ' While Loop
            
            If Eval(DecStr) Then
                WhileLoop.Add NumCurrentLine + 1 & ":" & DecStr
                
            Else
                Skip = True
            End If
            GoTo SkipThisSegment:
        Case 17   'For Block
            forT = DecStr
            forI = InStr(1, forT, "to", vbTextCompare)
            forTo = Trim(Mid(forT, 1, forI - 1)) ' for ile to arasý
            forTo2 = Trim(Mid(forT, forI + 2)) 'to sonrasý
            
            ChangeVar Mid(forTo, 1, InStr(1, forTo, "=") - 1), Mid(forTo, InStr(1, forTo, "=") + 1) ' for i=1
            
            
            forTo2 = Trim(Replace(forTo2, "=", "<>"))
            
            If Eval(forTo2) = True Then
                WhileLoop.Add NumCurrentLine + 1 & ":" & forT
                ChangeVar Mid(forTo, 1, InStr(1, forTo, "=") - 1), GetVar(Mid(forTo, 1, InStr(1, forTo, "=") - 1)) + 1  ' for i=1
            Else
                ForSkip = True
            End If
            GoTo SkipThisSegment:
            
        Case 18 ' Assignment
            tmpEq = Mid(DecStr, InStr(1, DecStr, "=") + 1)
            tmpAss = Trim(Mid(DecStr, 1, InStr(1, DecStr, "=") - 1))
            
            'Check if its one of our functions,See Public Sub MyF()
            'For o = 0 To UBound(MyFuncs)
                
                'j = InStr(1, tmpEq, MyFuncs(o), vbTextCompare)
                'If j <> 0 And InStr(1, tmpEq, "(") > 0 Then
                '    DebugAdd "CHANGE [" & tmpAss & "] BY FUNCTION [" & tmpEq & "]"
                '    ChangeVar tmpAss, DoFunction(Trim(tmpEq))
                '    GoTo SkipThisSegment:
                'End If
            'Next o
            DebugAdd "CHANGE [" & tmpAss & "] BY EXPRESSION [" & tmpEq & "]"
            ChangeVar tmpAss, evalY(Trim(tmpEq))
            GoTo SkipThisSegment:
            
        Case 19 ' Procedure Call
            Fok = 0
            For o = 0 To UBound(MyFuncs)
                If InStr(1, DecStr, MyFuncs(o), vbTextCompare) <> 0 Then Fok = 1
            Next o
            'If Fok = 0 Then MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Not a valid function : " + Mid(StrCurrentLine, 1, InStr(1, StrCurrentLine, "(") - 1) + vbCr + "This function does not exists in any library included.", vbExclamation: Exit Sub
            DebugAdd "CALL FUNCTION:[" & DecStr & "]"
            DoFunction DecStr
            GoTo SkipThisSegment:
        End Select
        
        
        
SkipThisSegment:
     
        
       
        DoEvents
       
    Next NumCurrentLine
    'Close #1
    'frmScript.Caption = Round((et - st) / 1000, 3) & " ms."
    'Show how much time had passed for execution ,silly isn't it?
    Close #2
    Erase Vars
    ExitPro
    Exit Sub
    'Close #OpenIt
erc:
    If Err.Number = 9 Then
        If CurVar > 63 Then
            MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Variable Declaration error!" + vbCr + "You can't use more than 64 local variables!", 16
        ElseIf CurrentGV > 255 Then
            MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Variable Declaration error!" + vbCr + "You can't use more than 255 global variables!", 16
        End If
    Else
        MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Unknown Error ", 16
    End If
End Sub
Private Sub CountP(ChkStr As String)
    'Check if Quotes and/or Parantheses match
    
    Dim t As Integer
    Dim t2 As Integer
    Dim q As Integer
    t = Len(ChkStr) - Len(Replace(ChkStr, "(", "", , Len(ChkStr)))
    t2 = Len(ChkStr) - Len(Replace(ChkStr, ")", "", , Len(ChkStr)))
    q = Len(ChkStr) - Len(Replace(ChkStr, Chr(34), "", , Len(ChkStr)))
    If q Mod 2 <> 0 Then MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Unmatching Quotes", 16: ECount = ECount + 1
    If t <> t2 Then MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Unmatching Parantheses", 16: ECount = ECount + 1
End Sub
Private Sub ChangeVar(TheVar As String, NewVal As Variant)
    'Change the value of a local,global or array variable
    
    Dim i As Integer
    Dim k As Integer
    For i = 0 To 63
        If Vars(i).Name = TheVar Then
            
            Vars(i).value = NewVal
            
            Exit Sub
        End If
    Next i
    'Change global
    ClientWindow.ChangeGlob TheVar, NewVal
    'Change array or system variable
    ch_Ident TheVar, NewVal
   
End Sub

Private Function GetVar(TheVar) As Variant
    'Get value of a variable
    Dim i As Integer
    Dim k As Integer
    DebugAdd "GETVAR :[" & TheVar & "]"
    For i = 0 To 63
        If Vars(i).Name = TheVar Then
            GetVar = Vars(i).value
            Exit Function
        End If
    Next i
    
    If ClientWindow.GetGlob(TheVar) <> "45lbb" Then
        
        GetVar = ClientWindow.GetGlob(TheVar)
        Exit Function
    End If
    
    GetVar = GetIdent(TheVar)
    
    DebugAdd "FINAL GETVAR:[" & GetVar & "]"
    'msgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Undeclared variable: " & TheVar, 16: BreakMe = True
End Function
Private Function GetIdent(Idcall) As Variant
    'Get Array,or system variables such as time,date app_path etc.
    Dim a1 As Byte
    Dim a2 As Byte
    Dim arr As Integer
    Dim nCall As String
    Dim AIndex As Integer
    Dim k As Integer
    Dim MainP As String
    Dim SubP As String
    Dim i As Integer
    Dim cnt As Byte
    
    'MsgBox "GET IDENT :[" & Idcall & "]"
    
    
    If InStr(1, Idcall, "[") > 0 Then
        AIndex = Val(evalY(Mid(Idcall, InStr(1, Idcall, "[") + 1, InStr(1, Idcall, "]") - InStr(1, Idcall, "[") - 1)))
        
        
        'GetIdent = ClientWindow.GetArr(Mid(Idcall, 1, InStr(1, Idcall, "[") - 1), AIndex)
        'Select Case Mid(Idcall, 1, InStr(1, Idcall, "[") - 1)
        '    Case "x"
        '
        'End Select
        Exit Function
        
    ElseIf InStr(1, Idcall, ".") > 0 Then
        MainP = LCase(Mid(Idcall, 1, InStr(1, Idcall, ".") - 1))
        SubP = LCase(Mid(Idcall, InStr(1, Idcall, ".") + 1, Len(Idcall) - InStr(1, Idcall, ".")))
        
        
    End If
    
    
    
found:
    
    
    
    
    
    nCall = Idcall
    'MsgBox "Idcall " & Trim(LCase(nCall))
    Select Case Trim(LCase(nCall))
        
    Case "time"
        GetIdent = Time
        Exit Function
    Case "cmd$"
        GetIdent = CmdL
        Exit Function
    Case "date"
        GetIdent = Date
        Exit Function
    Case "app_path"
        GetIdent = App.Path
        Exit Function
    Case "mousex"
        GetIdent = MouseX
        Exit Function
    Case "mousey"
        GetIdent = MouseY
        Exit Function
    Case "mbutton"
        GetIdent = Mbutton
        Exit Function
    Case "red"
        GetIdent = vbRed
        Exit Function
    Case "blue"
        GetIdent = vbBlue
        Exit Function
    Case "black"
        GetIdent = vbBlack
        Exit Function
    Case "white"
        GetIdent = vbWhite
        Exit Function
    Case "yellow"
        GetIdent = vbYellow
        Exit Function
    Case "gray"
        GetIdent = &H8000000F
        Exit Function
    Case "green"
        GetIdent = vbGreen
        Exit Function
    Case "crlf"
        GetIdent = vbCrLf
        Exit Function
        
    End Select
    
directEXit:
    DebugAdd "GET IDENT RETURN:[" & Idcall & "]"
    GetIdent = Idcall
End Function
Private Function Eval(sFormula As String) As Boolean
    'Basic Boolean Evaluation
    'if x<5 or s=0
    'something like these
    '================================
    
    DebugAdd "EVALUATE STATEMENT :[" & sFormula & "]"
    'On Error GoTo erEval
    Dim i As Integer
    Dim LeftVal As String, RightVal As String, Operator As String
    Dim sTemp As String
    Dim iq As Boolean
    Dim ib As Boolean
    
    ' Get the left value
    For i = 1 To Len(sFormula)
        sTemp = Mid(sFormula, i, 1)
        If sTemp = Chr(34) And iq = False Then
            iq = True
        ElseIf sTemp = "(" And ib = False And iq = False Then
            ib = True
        ElseIf sTemp = ")" And ib = True And iq = False Then
            ib = False
        ElseIf sTemp = Chr(34) And iq = True Then
            iq = False
        ElseIf IsOp(sTemp) = True And iq = False And ib = False Then
            If IsOp(Mid(sFormula, i + 1, 1)) = True Then
                Operator = sTemp & Mid(sFormula, i + 1, 1)
                LeftVal = Mid(sFormula, 1, i - 1)
                RightVal = Mid(sFormula, i + 2)
                
            ElseIf IsOp(Mid(sFormula, i + 1, 1)) = False Then
                Operator = sTemp
                LeftVal = Mid(sFormula, 1, i - 1)
                RightVal = Mid(sFormula, i + 1)
            End If
            Exit For
        End If
    Next i
    'MsgBox LeftVal & vbCr & Operator & vbCr & RightVal: Exit Function
    
    LeftVal = evalY(LeftVal)
    RightVal = evalY(RightVal)
    
    Select Case Operator
    Case ">"
        If Val(LeftVal) > Val(RightVal) Then Eval = True
    Case "<"
        If Val(LeftVal) < Val(RightVal) Then Eval = True
    Case "<>"
        If Val(LeftVal) <> Val(RightVal) Then Eval = True
    Case ">="
        If Val(LeftVal) >= Val(RightVal) Then Eval = True
    Case "<="
        If Val(LeftVal) <= Val(RightVal) Then Eval = True
    Case "="
        If LeftVal = RightVal Then Eval = True
    Case "=="
        If LCase(LeftVal) = LCase(RightVal) Then Eval = True
    End Select
    DebugAdd "EVALUATE [" & sFormula & "] RESULT [" & Eval & "]"
    
    Exit Function
erEval:
    MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Comparison Error " & sFormula, 16
End Function


Private Function IsOp(OPstr As String) As Boolean
    'Simple,shows if Opstr is <,> or =
    IsOp = False
    Select Case OPstr
    Case "=", "<", ">"
        IsOp = True
    End Select
    '3+5*2
End Function
Private Function Oops(st As String) As Boolean
    'Used in byte-compiling,see EncStr
    Oops = False
    Select Case st
    Case "+", "-", "*", "/", "=", "<", ">", "@", "(", ")", "&", Chr(34), ",", "^"
        Oops = True
        
    End Select
End Function
Private Function DoFunction(daCall As String) As Variant
    '===================================
    'Parse a function with its parameters
    '
    '===================================
    'On Error GoTo erdo
    Dim sName As String
    Dim arg() As String
    Dim i As Integer
    Dim argStr As String
    Dim iq As Boolean ' isQuote :)
    Dim argc As String
    Dim cchar As String * 1
    Dim oldp As Integer
    Dim p As Integer
    Dim p1 As Integer
    Dim p2 As Integer
    
    p1 = InStr(1, daCall, "(")
    p2 = InStrRev(daCall, ")")
    If p1 = 0 And p2 = 0 Then DoFunction = evalY(daCall): Exit Function
    
    sName = Trim(Mid(daCall, 1, p1 - 1))
    
    
    argStr = Mid(daCall, p1 + 1, p2 - p1 - 1)
    
    'MsgBox daCall + vbCr + sName + " => " + argStr
    
    
    iq = False
    
    For i = 1 To Len(argStr)
loophere:
        cchar = Mid(argStr, i, 1)
        If iq = False And cchar = "(" Then
            i = InStr(i, argStr, ")")
            GoTo loophere
        ElseIf iq = False And cchar = Chr(34) Then
            iq = True
        ElseIf iq = True And cchar = Chr(34) Then
            iq = False
        ElseIf iq = False And cchar = "," Then
            argc = argc + Mid(argStr, oldp + 1, i - oldp - 1) + ";"
            oldp = i
        End If
    Next i
    argc = argc + Mid(argStr, oldp + 1, Len(argStr) - oldp) + ";"
    arg = Split(argc, ";")
    ReDim Preserve arg(UBound(arg) - 1)
    For p = 0 To UBound(arg)
        arg(p) = evalY(arg(p))
    Next
    DoFunction = ExecuteFunction(sName, arg())
    Exit Function
erdo:
    
    MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Function Parse Error " & daCall, 16
End Function

Private Function ExecuteFunction(NameOfFunction As String, arg() As String) As Variant
'Execute a function,called by DoFunction(see above)
'Receives a func. name and array of params
'Runs included ActiveX if function can't be found here
'see Case else
    On Error GoTo ErrorEx:
    Dim i As Currency
    Dim user As Integer
    Dim u As Byte
    Dim t As Integer
    Dim tmx As Integer
    Dim tmy As Integer
    Dim TalkLen As Integer
    Dim oldXi As Integer
    Dim OldYi As Integer
    Dim tm As String
    
    For t = 0 To UBound(arg)
        tm = tm + arg(t)
    Next
    
    
    DebugAdd "EXECUTE FUNCTION:[" & NameOfFunction & "] WITH [" & (UBound(arg) + 1) & "] ARGUMENTS"
    Select Case LCase(NameOfFunction)
        
    Case "inc"
        If Left(arg(0), 1) <> "@" Then
            MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Error Executing function: " & NameOfFunction & vbCr & "@ character required to access address of a variable!", 16: Exit Function
        End If
        
        ChangeVar Mid(arg(0), 2, Len(arg(0)) - 1), Val(GetVar(Mid(arg(0), 2, Len(arg(0)) - 1))) + 1
        'MsgBox "Inside-" & GetVar(arg(0)) & vbCr & GetVar("levent")
    Case "dec"
        If Left(arg(0), 1) <> "@" Then
            MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Error Executing function: " & NameOfFunction & vbCr & "@ character required to access address of a variable!", 16: Exit Function
        End If
        
        ChangeVar Mid(arg(0), 2, Len(arg(0)) - 1), Val(GetVar(Mid(arg(0), 2, Len(arg(0)) - 1))) - 1
        'MsgBox "Inside-" & GetVar(arg(0)) & vbCr & GetVar("levent")
        
    Case "print"
        ClientWindow.Print arg(0)
        If ClientWindow.CurrentY > ClientWindow.Height Then
            ClientWindow.Height = ClientWindow.Height + 250
        End If
    Case "cls"
        ClientWindow.Cls
        
    Case "wait"
        Wait arg(0)
    Case "show"
        ClientWindow.Show
    Case "msgbox"
        If UBound(arg) = 0 Then MsgBox arg(0), 0, FileName(ScriptFile)
        If UBound(arg) = 1 Then ExecuteFunction = MsgBox(arg(0), Val(arg(1)), FileName(ScriptFile))
        If UBound(arg) = 2 Then ExecuteFunction = MsgBox(arg(0), Val(arg(1)), arg(2))
    Case "inputbox"
        If UBound(arg) = 0 Then ExecuteFunction = InputBox(arg(0), FileName(ScriptFile))
        If UBound(arg) = 1 Then ExecuteFunction = InputBox(arg(0), arg(1))
        If UBound(arg) = 2 Then ExecuteFunction = InputBox(arg(0), arg(1), arg(2))
    Case "end"
        BreakScript = True
        Unload Form1 'PostMessage ClientWindow.hWnd, &H10, 0, 0
    Case "mid"
        ExecuteFunction = Mid(arg(0), arg(1), arg(2))
    Case "instr"
        ExecuteFunction = InStr(arg(0), arg(1), arg(2))
    Case "instrrev"
        ExecuteFunction = InStrRev(arg(0), arg(1))
    Case "left"
        ExecuteFunction = Left(arg(0), arg(1))
    Case "right"
        ExecuteFunction = Right(arg(0), arg(1))
    Case "int"
        ExecuteFunction = Int(arg(0))
    Case "time"
        ExecuteFunction = Time
    Case "date"
        ExecuteFunction = Date
    Case "tick"
        ExecuteFunction = GetTickCount&
    Case "playsound"
        sndPlaySound arg(0), Val(arg(1))
    Case "random"
        Randomize Timer
        ExecuteFunction = Rnd
    Case "sin"
        
        ExecuteFunction = Sin(CSng(arg(0)))
    Case "cos"
        ExecuteFunction = Cos(CSng(arg(0)))
    Case "tan"
        ExecuteFunction = Tan(CSng(arg(0)))
    Case "ln"
        ExecuteFunction = Log(CSng(arg(0)))
    Case "abs"
        ExecuteFunction = Abs(CSng(arg(0)))
    Case "sqr"
        ExecuteFunction = Int(Sqr(Val(arg(0))))
        
    Case Else
        'Run external lib
        DebugAdd "EXTERNAL LIBRARY FUNCTION:[" & NameOfFunction & "]"
        For u = 0 To ClientWindow.Inclist.Count - 1
            ExecuteFunction = Incs(u).ExecuteFunction(NameOfFunction, arg())
        Next
        
        'MsgBox "else f " & NameOfFunction
        
        
        
    End Select
    'MsgBox "Execute : " + NameOfFunction + "(" + tm + ")" + vbCr & ExecuteFunction
    
    Exit Function
    
ErrorEx:
    If Err.Number <> 91 Then MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Error Executing function: " & NameOfFunction, 16
    Resume Next
End Function

Public Function evalY(expr As String)
    '---------------------------
    'Main evaluation sub
    'Evaluates anything,such as
    'a=4+5,or msgbox("levent " & "rocks")
    'Pretty buggy,be careful when modifying.
    '---------------------------
    On Error GoTo erev
    Dim value As Variant, operand As String
    Dim pos As Integer
    Dim fxx As Integer
    
    
    pos = 1
    'Print #OpenIt, (NumCurrentLine + 1) & " " & expr
    'MsgBox "EVALY [" & expr & "]"
    Do Until pos > Len(expr)
        'SkipQU = 1
        
        Select Case Mid(expr, pos, 3)
        Case "not", "or ", "and", "xor", "eqv", "imp", "mod"
            operand = Mid(expr, pos, 3)
            pos = pos + 3
        End Select
        
        
        Select Case Mid(expr, pos, 1)
        Case " "
            pos = pos + 1
        Case "&", "+", "-", "*", "/", "\", "^"
            operand = Mid(expr, pos, 1)
            
            pos = pos + 1
            
        Case ">", "<", "=":
            
            
            Select Case Mid(expr, pos + 1, 1)
            Case "<", ">", "="
                operand = Mid(expr, pos, 2)
                pos = pos + 1
            Case Else
                operand = Mid(expr, pos, 1)
            End Select
            pos = pos + 1
        Case Else
            If Mid(expr, pos, 1) = Chr(34) Then
                If SkipQU = 0 Then
                    'MsgBox "VALUE BEFORE QUOTE:[" & value & "]"
                    value = value + Mid(expr, pos + 1, InStr(pos + 1, expr, Chr(34)) - pos - 1)
                    'MsgBox "VALUE AFTER  QUOTE:[" & value & "]"
                    pos = pos + InStr(pos + 1, expr, Chr(34)) + 1
                    'MsgBox operand, 0, Len(operand)
                    
                End If
            End If
            
            'MsgBox "OPERAND:[" & operand & "]"
            Select Case operand
            Case "": evalY = Tokenize(expr, pos)
            Case "&": evalY = evalY & value
                value = Tokenize(expr, pos)
            Case "+": evalY = evalY + value
                value = Tokenize(expr, pos)
            Case "-": evalY = evalY + value
                value = -Tokenize(expr, pos)
                
            Case "*": value = value * Tokenize(expr, pos)
            Case "/": value = value / Tokenize(expr, pos)
            Case "\": value = value \ Tokenize(expr, pos)
            Case "^": value = value ^ Tokenize(expr, pos)
            Case "not": evalY = evalY + value
                value = Not Tokenize(expr, pos)
            Case "and": value = value And Tokenize(expr, pos)
            Case "or ": value = value Or Tokenize(expr, pos)
            Case "mod": value = value Mod Tokenize(expr, pos) ': MsgBox Value, 0, Tokenize(expr, pos)
            Case "xor": value = value Xor Tokenize(expr, pos)
            Case "eqv": value = value Eqv Tokenize(expr, pos)
            Case "imp": value = value Imp Tokenize(expr, pos)
            Case "=", "==": value = value = Tokenize(expr, pos)
            Case ">": value = value > Tokenize(expr, pos)
            Case "<": value = value < Tokenize(expr, pos)
            Case ">=", "=>": value = value >= Tokenize(expr, pos)
            Case "<=", "=<": value = value <= Tokenize(expr, pos)
            Case "<>": value = value <> Tokenize(expr, pos)
            End Select
        End Select
    Loop
    'MsgBox "EVALY VALUE :[" & value & "]"
    If Not IsNumeric(value) Then
        
        evalY = evalY + value
    ElseIf IsNumeric(value) And value <> 0 Then
        evalY = evalY + Str(value)
    End If
    
    'MsgBox "RESULT EVALY :[" & expr & "] TO :[" & evalY & "]"
    
    Exit Function
erev:
    MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Error Evaluating Statement " & expr, 16
    
End Function


Private Function Tokenize(expr, pos)
    '----------------------------
    'Here we tokenize Expressions
    'to evaluate them
    'called by EvalY,and calls evalY.
    'Pretty buggy,be careful when modifying.
    '----------------------------
    'On Error GoTo ertok
    Dim char As String, value As String, fn As String
    Dim es As Integer, pl As Integer
    Dim fc As Integer
    Dim SHR As String
    
    Const QUOTE As String = """"
    
    SkipQU = 1
    'MsgBox "TOKEN THIS [" & expr & "] POS [" & pos & "]"
    'Print #OpenIt, (NumCurrentLine + 1) & " " & pos & " " & expr
    SHR = expr
    
    'If pos > Len(expr) Then Tokenize = evalY(SHR)
    
    Do Until pos > Len(expr)
        char = Mid(expr, pos, 1)
        
        
        
        Select Case char
        Case "&", "+", "-", "/", "\", "*", "^", ">", "<", "=": DebugAdd "TOKENIZE OPERATOR IN [" & expr & "] POS [" & "]": Exit Do
            
        Case "("
            pl = 1
            pos = pos + 1
            es = pos
            'DebugAdd "TOKENIZE PAR '(' IN [" & expr & "] POS [" & "]"
            
            Do Until pl = 0 Or pos > Len(expr)
                char = Mid(expr, pos, 1)
                
                
                Select Case char
                Case "(": pl = pl + 1
                Case ")": pl = pl - 1
                End Select
                'MsgBox pos & " " & pl & " " & expr
                pos = pos + 1
            Loop
            value = Mid(expr, es, pos - es - 1)
            fn = LCase(Tokenize)
            
            For fc = 0 To UBound(MyFuncs)
                If fn = MyFuncs(fc) Then
                    'DebugAdd "TOKEN DOFUNC FUNCNAME:[" & fn & "]"
                    
                    Tokenize = DoFunction(fn + "(" + value + ")")
                    'MsgBox fn + "[" + value + "]", 0, Tokenize
                    'DebugAdd "TOKEN DOFUNC RETURN  :[" & Tokenize & "]"
                    
                    GoTo SEXIT
                End If
            Next
            'DebugAdd "TOKEN NOFUNC VALUE   :[" & Value & "]"
            
            Tokenize = evalY(value)
            
            
            Exit Do
        Case QUOTE
            pl = 1
            'pos = InStr(pos, expr, Chr(34))
            pos = pos + 1
            es = pos
            
            'DebugAdd "TOKENIZE QUOTE IN [" & expr & "] POS [" & pos & "]"
            Do Until pl = 0 Or pos > Len(expr)
                char = Mid(expr, pos, 1)
                pos = pos + 1
                
                
                If char = QUOTE Then
                    
                    If Mid(expr, pos, 1) = QUOTE Then
                        value = value & QUOTE
                        
                        pos = pos + 1
                    Else
                        Exit Do
                    End If
                    
                Else
                    value = value & char
                End If
                
            Loop
            
            Tokenize = value
            Exit Do
        Case Else
            'DebugAdd "TOKENIZE ELSE [" & expr & "] POS [" & pos & "]"
            Tokenize = Tokenize & char
            pos = pos + 1
        End Select
    Loop
    
SEXIT:
    'MsgBox "LAST TOKENIZE [" & Tokenize & "]"
    If IsNumeric(Tokenize) And Tokenize <> "" Then
        'DebugAdd "$TOKENIZE NUMERIC"
        Tokenize = Val(Tokenize)
    ElseIf Left(Tokenize, 1) = "@" Then
        'DebugAdd "$TOKENIZE VARIABLE ADDRESS"
    ElseIf IsDate(Tokenize) Then
        'DebugAdd "$TOKENIZE DATE"
        Tokenize = CDate(Tokenize)
    Else
        'DebugAdd "$TOKENIZE GETVAR"
        Tokenize = GetVar(Tokenize)
    End If
    
    'MsgBox "FINAL TOKENIZE VALUE :[" & Tokenize & "]"
    Exit Function
ertok:
    MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + "Script Tokenization Error " & expr, 16
    
End Function
Private Sub DebugAdd(Node As String)
    'Slows down code I don't use it anymore
    'switch on if you like
    'If DMode = True Then DebugW.List1.AddItem "ln[" & (NumCurrentLine + 1) & "]" + Node
End Sub
Private Sub ch_Ident(ID As String, newV)
    '------------------------
    'Change system variables or arrays
    '------------------------
    Dim AIndex As Integer
    Dim i As Integer
    Dim cnt As Byte
    'DebugAdd "CHANGE IDENT :[" & ID & "] TO [" & newV & "]"
    Dim MainP As String
    Dim SubP As String
    If InStr(1, ID, "[") > 0 Then
        AIndex = Val(evalY(Mid(ID, InStr(1, ID, "[") + 1, InStr(1, ID, "]") - InStr(1, ID, "[") - 1)))
        
        ClientWindow.ChangeArr Mid(ID, 1, InStr(1, ID, "[") - 1), AIndex, newV
        'Select Case Mid(ID, 1, InStr(1, ID, "[") - 1)
        '    Case "x"
        ' Sprites(AIndex).x = newV
        
        'End Select
        Exit Sub
        
    ElseIf InStr(1, ID, ".") > 0 Then
        MainP = LCase(Mid(ID, 1, InStr(1, ID, ".") - 1))
        SubP = LCase(Mid(ID, InStr(1, ID, ".") + 1, Len(ID) - InStr(1, ID, ".")))
        
    Else
        Select Case ID
            'just some old game params,i was coding
            'nothing to do
        Case "room_name"
            
        Case "armed"
            
        End Select
    End If
    
    'MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error Line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + ID & " is a Sytem_Ident_Variable!" + vbCr + "You can't assign values to them!", 16
    
End Sub
Private Sub Exception(Expl As String, Fatal As Byte)
    'Nothing special here
    MsgBox "Error Module [" + ScriptFile + "]" + vbCr + "Error Procedure [" + Procedure + "]" + vbCr + "Error Line [" & (NumCurrentLine + 1) & "]" + vbCr + vbCr + Expl, 16
    
    
End Sub

Public Sub Compile()
    '------------------------------------
    'Sub that compiles a plain txt file to
    'byte code file
    '------------------------------------
    Dim ln() As String
    
    Dim tmLast As String
    Dim iTemp As Integer
    Dim i As Integer
    Dim Nline As Integer
    
    Dim tmpGlob As String
    Dim GoToLabel As String
    
    
    
    'Comment scanning vars
    Dim aL As Integer
    Dim aL2 As Integer
    Dim bL As Integer
    Dim bL2 As Integer
    Dim Comment As String
    '-------------------
    
    
    Do
        aL = InStr(1 + aL, Script, "/*")
        If aL = 0 Then Exit Do
        aL2 = InStr(1 + aL, Script, "*/")
        Comment = Mid(Script, aL, aL2 - aL + 2)
        Script = Replace(Script, Comment, "") 'vb's advantage :)
    Loop
    
    ln = Split(Script, ";")
    Nline = UBound(ln) 'number of lines
    
    
    
    
    
    Put #1, , Nline
    For n = 0 To UBound(ln)
        If Left(ln(n), 2) = "//" Then GoTo SkipThisSegment:
        
        CForm.Label1.Caption = "Compiling [" & FileName(DestFile) & "] " & ln(n)
        iTemp = InStr(1, ln(n), " ")
        CountP ln(n)
        
        If ln(n) = GoToLabel Then
            Parser.Comm = 1
            EncStr (GoToLabel)
            GoToLabel = ""
            GoTo SkipThisSegment:
        End If
        
        
        
        
        If LCase(ln(n)) = "wend" Then
            Parser.Comm = 2
            
            GoTo SkipThisSegment:
        End If
        
        
        
        
        If LCase(ln(n)) = "next" Then
            Parser.Comm = 3
            
            GoTo SkipThisSegment:
        End If
        
        
        
        
        Select Case LCase(ln(n))
        Case "endif"
            Parser.Comm = 4
            GoTo SkipThisSegment:
        Case "else"
            Parser.Comm = 5
            
            GoTo SkipThisSegment:
        End Select
        
        
        
        
        ' Determine what is the rest
        iTemp = InStr(1, ln(n), " ")
        If iTemp <> 0 Then
            Select Case LCase(Mid(ln(n), 1, iTemp - 1))
                
            Case "if" 'if block
                Parser.Comm = 9
                EncStr (Mid(ln(n), iTemp + 1))
                GoTo SkipThisSegment:
                
            Case "goto" ' Goto a label
                Parser.Comm = 10
                EncStr (Trim(Mid(ln(n), iTemp + 1)))
                
                GoTo SkipThisSegment:
            Case "global" 'Declare a global variable
                Parser.Comm = 11
                EncStr (Mid(ln(n), iTemp + 1))
                
                
                
                GoTo SkipThisSegment:
            Case "array"
                Parser.Comm = 12
                
                EncStr (Mid(ln(n), iTemp + 1))
                
                'ClientWindow.AddArray Trim(tmpGlob)
                GoTo SkipThisSegment:
                
            Case "include"
                Parser.Comm = 13
                EncStr (Mid(ln(n), iTemp + 1))
                
                
            Case "int" ' Declare an integer variable
                Parser.Comm = 14
                EncStr (Mid(ln(n), iTemp + 1))
                'CurVar = CurVar + 1
                GoTo SkipThisSegment:
            Case "string" ' Declare a string variable
                Parser.Comm = 15
                EncStr (Mid(ln(n), iTemp + 1))
                'CurVar = CurVar + 1
                
                GoTo SkipThisSegment:
            Case "while" ' While Loop
                Parser.Comm = 16
                EncStr (Mid(ln(n), iTemp + 1))
                
                GoTo SkipThisSegment:
            Case "for"   'For Block
                Parser.Comm = 17
                EncStr (Mid(ln(n), iTemp + 1))
                GoTo SkipThisSegment:
            End Select
            
        End If
        'end routines
        For i = 1 To Len(ln(n))
            Select Case Mid(ln(n), i, 1)
            Case "=" ' Assignment
                Parser.Comm = 18
                
                Dim tmpEq As String
                Dim tmpAss As String
                
                tmpEq = Mid(ln(n), InStr(1, ln(n), "=") + 1)
                tmpAss = Trim(Mid(ln(n), 1, InStr(1, ln(n), "=") - 1))
                
                If CForm.Check1.value = 1 And IsNumeric(optimize(tmpEq)) = True Then
                'Optimization see below
                EncStr (tmpAss & "=" & optimize(tmpEq))
                Else
                EncStr (ln(n))
                End If
                
                
                GoTo SkipThisSegment:
                
            Case "(" ' Procedure Call
                Parser.Comm = 19
                
                If CForm.Check1.value = 1 Then
                '-------------------------
                'Optimize args of function
                'E.g msgbox(12+5,0,"hello")
                'optimized as:
                'msgbox(17,0,"hello")
                '-------------------------
                EncStr (OptFunction(ln(n)))
                Else
                EncStr (ln(n))
                End If
                GoTo SkipThisSegment:
            End Select
            
        Next i
        
SkipThisSegment:
        'Do not put segment if its empty
        If Parser.Comm <> 0 Then Put #1, , Parser
        Parser.Comm = 0
        Erase Parser.Def
        Erase Parser.expr
        DoEvents
        Wait 0.01 'totally useless,only to see cool compiling outputs
    Next
    
    If ECount = 0 Then
        CForm.Label1.Caption = "Succesfully compiled [" & FileName(DestFile) & "]"
        'MsgBox "Compiled succesfully." + vbCr + "Input " & Len(Script) & " bytes" + vbCr + "Output " & LOF(1) & " bytes.", 64
    Else
        CForm.Label1.Caption = "Compiled [" & FileName(DestFile) & "] with " & ECount & " error(s)."
        'MsgBox "Compiled with " & ECount & " error(s)." + vbCr + "Input " & Len(Script) & " bytes" + vbCr + "Output " & LOF(1) & " bytes.", 16
    End If
    
End Sub
Private Function EncStr(StrI As String)
    'Encodes remaining string params or arguments
    'of functions and statements
    '
    Dim i As Integer
    Dim res As String
    If StrI = "" Then Exit Function
    ReDim Parser.expr(Len(StrI) - 1)
    ReDim Parser.Def(Len(StrI) - 1)
    

    
    For i = 0 To Len(StrI) - 1
        If Oops(Mid(StrI, i + 1, 1)) Then
            
            Parser.expr(i) = OpCode(Mid(StrI, i + 1, 1))
            Parser.Def(i) = 0 'operator
            
        ElseIf IsNumeric(Mid(StrI, i + 1, 1)) Then
            Parser.expr(i) = Val(Mid(StrI, i + 1, 1))
            Parser.Def(i) = 1 'num
        Else
            Parser.expr(i) = Asc(Mid(StrI, i + 1, 1))
            Parser.Def(i) = 2 'str
        End If
    Next
    
    
End Function
Private Function DecStr() As String
    'Reverse operation of EncStr
    'to run byte code.Called in RunBlock
    On Error Resume Next
    Dim i As Integer
    Dim res As String
    
    For i = 0 To UBound(Parse(NumCurrentLine).expr)
        If Parse(NumCurrentLine).Def(i) = 0 Then
            DecStr = DecStr + COpCode(Parse(NumCurrentLine).expr(i))
            
        ElseIf Parse(NumCurrentLine).Def(i) = 1 Then
            DecStr = DecStr + Trim(Str(Parse(NumCurrentLine).expr(i)))
            'num
        ElseIf Parse(NumCurrentLine).Def(i) = 2 Then
            DecStr = DecStr & Chr(Parse(NumCurrentLine).expr(i))
            
        End If
    Next
    'If InStr(DecStr, "jk") > 0 Then MsgBox DecStr, 0, "decStr"
    'MsgBox DecStr
End Function
Private Function OpCode(Op As String) As Byte
    'Used by EncStr,to convert Operators to byte-code
    Select Case Op
    Case "+"
        OpCode = 1
    Case "-"
        OpCode = 2
    Case "*"
        OpCode = 3
    Case "/"
        OpCode = 4
    Case "("
        OpCode = 8
    Case ")"
        OpCode = 9
    Case "&"
        OpCode = 11
    Case "^"
        OpCode = 12
        
        
    Case "<"
        OpCode = 5
    Case ">"
        OpCode = 7
    Case "="
        OpCode = 6
        
    Case Chr(34)
        OpCode = 10
    Case ","
        OpCode = 20
        
    End Select
End Function
Private Function COpCode(Op As Byte) As String
    'Reverse or COpcode
    Select Case Op
    Case 1
        COpCode = "+"
    Case 2
        COpCode = "-"
    Case 3
        COpCode = "*"
    Case 4
        COpCode = "/"
    Case 8
        COpCode = "("
    Case 9
        COpCode = ")"
    Case 11
        COpCode = "&"
    Case 12
        COpCode = "^"
        
    Case 5
        COpCode = "<"
    Case 7
        COpCode = ">"
    Case 6
        COpCode = "="
        
    Case 10
        COpCode = Chr(34)
    Case 20
        COpCode = ","
        
    End Select
End Function
Private Sub ExitPro()
    'Exit applet,unregister dlls if any
    'Called at the end of RunBlock
    Dim i As Integer
    
    
    For i = 0 To ClientWindow.Inclist.Count - 1
        Shell "regsvr32.exe /s /u " & Chr(34) & ClientWindow.Inclist(i + 1) + ".dll" & Chr(34), vbHide
    Next
End Sub

