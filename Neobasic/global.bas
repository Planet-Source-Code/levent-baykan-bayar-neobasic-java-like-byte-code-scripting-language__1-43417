Attribute VB_Name = "global"
'Public Type GlobVar
'GName As String
'GVal As Variant
'End Type

'Public GVars(255) As GlobVar

'Type for byte-code segments

Public Type Segment
Comm As Byte 'defines func,operator,identifier type
Def() As Byte 'special info
expr() As Byte 'usually encoded info
End Type


Public CurrentGV As Integer
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetTickCount& Lib "kernel32" ()
Public MyFuncs() As String
Public MyData() As String

Sub MyF()
'Load function names into array
'used in evalY,Tokenize and RunBlock
Dim f As String
Dim b As String
b = "gotolabel;wend;next;endif;else;dump;dump;dump;if;goto;global;array;include;int;string;while;for;assign;proc call"

f = "inc;dec;msgbox;inputbox;wait;show;print;end;mid;instr;instrrev;left;right;int;time;date;tick;playsound;random;sin;cos;tan;abs;sqr;ln;cls"
MyData = Split(b, ";")

MyFuncs = Split(f, ";")

End Sub
Public Sub Wait(howlong)
'Wait by howlong
Dim temptime

temptime = Timer
Do
DoEvents
Loop While Timer < temptime + howlong
End Sub
Public Function FileName(WithPath As String)
'Retrieve filename of a path+filename string
'Ex: Filename("c:\levent.txt") => "levent.txt"
Dim sWithoutPath As String
Dim iLen As Integer
Dim iWhere As Integer
If InStr(1, WithPath, "\") = 0 Then FileName = WithPath: Exit Function
sWithoutPath = WithPath
Do Until InStr(sWithoutPath, "\") = 0
iLen = Len(sWithoutPath)
iWhere = InStr(sWithoutPath, "\")
sWithoutPath = Right(sWithoutPath, iLen - iWhere)
Loop
FileName = sWithoutPath

End Function
Public Function optimize(expr As String)

'Optimize math code
'derived from EvalY

On Error GoTo erop
    Dim value As Variant, operand As String
    Dim pos As Integer
    
    pos = 1


    Do Until pos > Len(expr)


        Select Case Mid(expr, pos, 3)
            Case "not", "or ", "and", "xor", "eqv", "imp"
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


    Select Case operand
        Case "": value = opToken(expr, pos)
        Case "&": optimize = optimize & value
        value = opToken(expr, pos)
        Case "+": optimize = optimize + value
        value = opToken(expr, pos)
        Case "-": optimize = optimize + value
        value = -opToken(expr, pos)
        
        Case "*": value = value * opToken(expr, pos)
        Case "/": value = value / opToken(expr, pos)
        Case "\": value = value \ opToken(expr, pos)
        Case "^": value = value ^ opToken(expr, pos)
        Case "not": optimize = optimize + value
        value = Not opToken(expr, pos)
        Case "and": value = value And opToken(expr, pos)
        Case "or ": value = value Or opToken(expr, pos)
        Case "xor": value = value Xor opToken(expr, pos)
        Case "eqv": value = value Eqv opToken(expr, pos)
        Case "imp": value = value Imp opToken(expr, pos)
        Case "=", "==": value = value = opToken(expr, pos)
        Case ">": value = value > opToken(expr, pos)
        Case "<": value = value < opToken(expr, pos)
        Case ">=", "=>": value = value >= opToken(expr, pos)
        Case "<=", "=<": value = value <= opToken(expr, pos)
        Case "<>": value = value <> opToken(expr, pos)
    End Select
End Select
Loop
If IsNumeric(optimize) = True And IsNumeric(value) = True Then
optimize = optimize + value
Else
optimize = optimize & value
End If
'MsgBox expr & vbCr & optimize & vbCr & Val(optimize) & vbCr & pos
Exit Function
erop:

optimize = expr
End Function


Private Function opToken(expr, pos)
'Optimize math system
'Derived from Tokenize

On Error GoTo erex
    Dim char As String, value As String, fn As String
    Dim es As Integer, pl As Integer
    Const QUOTE As String = """"
    


    Do Until pos > Len(expr)
        char = Mid(expr, pos, 1)


        Select Case char
            Case "&", "+", "-", "/", "\", "*", "^", " ", ">", "<", "=": Exit Do
            Case "("
            pl = 1
            pos = pos + 1
            es = pos


            Do Until pl = 0 Or pos > Len(expr)
                char = Mid(expr, pos, 1)


                Select Case char
                    Case "(": pl = pl + 1
                    Case ")": pl = pl - 1
                End Select
            pos = pos + 1
        Loop
        value = Mid(expr, es, pos - es - 1)
        fn = LCase(opToken)


        Select Case fn
            Case "sin": opToken = Sin(optimize(value))
            Case "cos": opToken = Cos(optimize(value))
            Case "tan": opToken = Tan(optimize(value))
            Case "exp": opToken = Exp(optimize(value))
            Case "log": opToken = Log(optimize(value))
            Case "atn": opToken = Atn(optimize(value))
            Case "abs": opToken = Abs(optimize(value))
            Case "sgn": opToken = Sgn(optimize(value))
            Case "sqr": opToken = Sqr(optimize(value))
            Case "rnd": opToken = Rnd(optimize(value))
            Case "int": opToken = Int(optimize(value))
            Case "day": opToken = Day(optimize(value))
            Case "month": opToken = Month(optimize(value))
            Case "year": opToken = Year(optimize(value))
            Case "weekday": opToken = Weekday(optimize(value))
            Case "hour": opToken = Hour(optimize(value))
            Case "minute": opToken = Minute(optimize(value))
            Case "second": opToken = Second(optimize(value))
            Case "date": opToken = Date
            Case "date$": opToken = Date$
            Case "time": opToken = Time
            Case "time$": opToken = Time$
            Case "timer": opToken = Timer
            Case "now": opToken = Now()
            Case "len": opToken = Len(optimize(value))
            Case "trim": opToken = Trim(optimize(value))
            Case "ltrim": opToken = LTrim(optimize(value))
            Case "rtrim": opToken = RTrim(optimize(value))
            Case "ucase": opToken = UCase(optimize(value))
            Case "lcase": opToken = LCase(optimize(value))
            Case "val": opToken = Val(optimize(value))
            Case "chr": opToken = Chr(optimize(value))
            Case "asc": opToken = Asc(optimize(value))
            Case "space": opToken = Space(optimize(value))
            Case "hex": opToken = Hex(optimize(value))
            Case "oct": opToken = Oct(optimize(value))
            Case Else: opToken = optimize(value)
        End Select
    Exit Do
    Case QUOTE
    pl = 1
    pos = pos + 1
    es = pos


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
    opToken = value
    Exit Do
    Case Else
    opToken = opToken & char
    pos = pos + 1
End Select
Loop



If IsNumeric(opToken) Then
opToken = Val(opToken)
ElseIf IsDate(opToken) Then
opToken = CDate(opToken)
End If
Exit Function
erex:
opToken = expr
End Function

Public Function OptFunction(daCall As String) As String
'Optimize functions args
'Derived from DoFunction

Dim sName As String
                Dim arg() As String
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
                sName = Trim(Mid(daCall, 1, p1 - 1))
                
                
                argStr = Mid(daCall, p1 + 1, p2 - p1 - 1)
                
                'MsgBox daCall + vbCr + sName + " => " + argStr
                
                
                iq = False
                
                For p = 1 To Len(argStr)
loophere:
                    cchar = Mid(argStr, p, 1)
                    If iq = False And cchar = "(" Then
                        p = InStr(p, argStr, ")")
                        GoTo loophere
                    ElseIf iq = False And cchar = Chr(34) Then
                        iq = True
                    ElseIf iq = True And cchar = Chr(34) Then
                        iq = False
                    ElseIf iq = False And cchar = "," Then
                        argc = argc + Mid(argStr, oldp + 1, p - oldp - 1) + ";"
                        oldp = p
                    End If
                Next p
                argc = argc + Mid(argStr, oldp + 1, Len(argStr) - oldp) + ";"
                arg = Split(argc, ";")
                ReDim Preserve arg(UBound(arg) - 1)
                For p = 0 To UBound(arg)
                    If IsNumeric(optimize(arg(p))) Then arg(p) = optimize(arg(p))
                Next
                argc = ""
                For p = 0 To UBound(arg)
                    argc = argc + arg(p) + ","
                Next
                argc = Left(argc, Len(argc) - 1)
                OptFunction = sName + "(" + argc + ")"
                
                Erase arg
                
    
End Function
