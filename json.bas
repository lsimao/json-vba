Attribute VB_Name = "JSON"
'JSON to VBA parser and stringifier
'2015-08-20, https://github.com/lsimao/json-vba/
'
'Mapping:
'=============================================
'| from JSON |       VB        |   to JSON   |
'=============================================
'| String    | String          | String      |
'+-----------+-----------------+-------------+
'| Number    | Double          | Number      |
'|           | Numeric (other) |             |
'|           | Date            |             |
'+-----------+-----------------+-------------+
'| Object    | Dictionary      | Object      |
'+-----------+-----------------+-------------+
'| Array     | Array           | Array       |
'|           | Collection      |             |
'+-----------+-----------------+-------------+
'| true      | True            | true        |
'| false     | False           | false       |
'+-----------+-----------------+-------------+
'| null      | Null            | null        |
'|           | Nothing         |             |
'|           | Empty           |             |
'|           | Missing         |             |
'+-----------+-----------------+-------------+
'|           | Class.toJSON    | (see Notes) |
'+-----------+-----------------+-------------+
'|           | Array 2D+       | Unsupported |
'|           | Object (other)  |             |
'|           | User defined    |             |
'|           | Class           |             |
'+-----------+-----------------+-------------+
'
'Notes:
'JSON Objects are mapped to VB Scripting.Dictionary objects and vice-versa.
'VB Date is casted to double and mapped to JSON Number.
'VB Collection is mapped to JSON Array.
'VB Objects/Classes excluding Dictionary and Collection are supported provided a toJSON member is available. toJSON return is not checked.
'User defined types and multidimensional arrays are not supported.
'Because JSON.Parse may return either an object or a value, JSON.Assign function is provided so you may easily assign function return to variable.
'
'Examples:
'
' 'Object assignment:
' Set v = JSON.Parse("{""a"":[0,1],""b"":100.2,""c"":null}")
' Debug.Print "b="; v!b
' Debug.Print "a[0]="; v!a(0)
' Debug.Print "Back to JSON="; JSON.Stringify(v)
'
' 'Value assignment:
' v = JSON.Parse("[0, ""abc""]")
' Debug.Print "Second element = "; v(1)
'
' 'Object or value assignement:
' JSON.Assign v, JSON.Parse(str)
'
Option Explicit

Public Const ERR_TRUNCATED As Long = 57001
Public Const ERR_UNEXPECTED_TOKEN As Long = 57002
Public Const ERR_EXPECTED_STRING As Long = 57003
Public Const ERR_INVALID_VALUE As Long = 57004
Public Const ERR_INVALID_ESCAPE As Long = 57005
Public Const ERR_UNSUPPORTED_TYPE As Long = 57011

Private Const ERR_SRC = "JSON"
Private Const ERR_TRUNCATED_DESC As String = "Data truncated at character %!"
Private Const ERR_UNEXPECTED_TOKEN_DESC As String = "Unexpected token at character %!"
Private Const ERR_EXPECTED_STRING_DESC As String = "Expected string at character %!"
Private Const ERR_INVALID_VALUE_DESC As String = "Invalid value at character %!"
Private Const ERR_INVALID_ESCAPE_DESC As String = "Invalid escape sequence at character %!"
Private Const ERR_UNSUPPORTED_TYPE_DESC As String = "Unsupported data type: $!"

'Front-end parser function
Public Function Parse(ByVal t As String)
    Dim v, p As Long
    p = 1
    Assign v, getValue(t, p)
    skipWS t, p
    If p < Len(t) Then Err.Raise ERR_UNEXPECTED_TOKEN, ERR_SRC, Replace(ERR_UNEXPECTED_TOKEN_DESC, "%", p)
    If IsObject(v) Then
        Set Parse = v
    Else
        Parse = v
    End If
End Function

'Recursive stringifier function
Public Function Stringify(d) As String
    Dim r As String, i As Long, c As Long, v
    'Support to objects implementing toJSON member
    On Error Resume Next
    Stringify = d.toJSON
    If Err.Number = 0 Then Exit Function
    On Error GoTo 0
    If (VarType(d) And vbArray) Or TypeName(d) = "Collection" Then
        If VarType(d) And vbArray Then
            On Error Resume Next 'Handle zero-length arrays
            i = LBound(d)
            c = i - 1
            c = UBound(d)
        Else
            i = 1
            c = d.Count
        End If
        For i = i To c
            r = IIf(Len(r), r & ",", vbNullString) & Stringify(d(i))
        Next
        r = "[" & r & "]"
    ElseIf VarType(d) = vbString Then
        For i = 1 To Len(d)
            c = AscW(Mid(d, i, 1))
            Select Case c
            Case &H8: r = r & "\b" 'Backspace
            Case &H9: r = r & "\t" 'Tab
            Case &HA: r = r & "\n" 'Line feed
            Case &HC: r = r & "\f" 'Form feed
            Case &HD: r = r & "\r" 'Carriage return
            Case &H22: r = r & "\""" 'Quotation mark
            Case &H2F: r = r & "\/" 'Solidus character
            Case &H5C: r = r & "\\" 'Reverse solidus character
            Case 0 To &H1F, &H7F 'Control characters
                r = r & "\u" & Right("000" & Hex(c), 4)
            Case Else
                r = r & ChrW(c)
            End Select
        Next
        r = """" & r & """"
    ElseIf VarType(d) = vbBoolean Then
        r = IIf(d, "true", "false")
    ElseIf IsNumeric(d) Then
        r = Replace(d, Mid(1.3, 2, 1), ".")
    ElseIf IsDate(d) Then
        r = Replace(CDbl(d), Mid(1.3, 2, 1), ".")
    ElseIf IsNull(d) Or IsEmpty(d) Or IsMissing(d) Then
        r = "null"
    ElseIf TypeName(d) = "Dictionary" Then
        v = d.keys
        For i = LBound(v) To UBound(v)
            r = IIf(i > LBound(v), r & ",", vbNullString) & Stringify(v(i)) & ":" & Stringify(d(v(i)))
        Next
        r = "{" & r & "}"
    ElseIf IsObject(d) Then
        If d Is Nothing Then
            r = "null"
        Else
            Err.Raise ERR_UNSUPPORTED_TYPE, ERR_SRC, Replace(ERR_UNSUPPORTED_TYPE_DESC, "$", TypeName(d))
        End If
    Else
        Err.Raise ERR_UNSUPPORTED_TYPE, ERR_SRC, Replace(ERR_UNSUPPORTED_TYPE_DESC, "$", TypeName(d))
    End If
    Stringify = r
End Function

'Helper function, used to assign either objects or values in a single statement
Public Sub Assign(ByRef Target As Variant, Source)
    If IsObject(Source) Then
        Set Target = Source
    Else
        Target = Source
    End If
End Sub

'Worker recursive function to parse JSON
Private Function getValue(t As String, p As Long)
    Dim v, i As Long, k
    skipWS t, p
    Select Case Mid(t, p, 1)
    Case "{"
        Set v = CreateObject("Scripting.Dictionary")
        v.CompareMode = vbBinaryCompare
        skipWS t, p, 1
        If Mid(t, p, 1) <> "}" Then
            Do
                If Mid(t, p, 1) <> """" Then Err.Raise ERR_EXPECTED_STRING, ERR_SRC, Replace(ERR_EXPECTED_STRING_DESC, "%", p)
                Assign k, getValue(t, p)
                skipWS t, p
                If Mid(t, p, 1) <> ":" Then Err.Raise ERR_UNEXPECTED_TOKEN, ERR_SRC, Replace(ERR_UNEXPECTED_TOKEN_DESC, "%", p)
                skipWS t, p, 1
                If v.exists(k) Then v.Remove k
                v.Add k, getValue(t, p)
                skipWS t, p
                Select Case Mid(t, p, 1)
                Case "}" 'Last pair
                    p = p + 1
                    Exit Do
                Case ","
                    skipWS t, p, 1
                Case Else
                    Err.Raise ERR_UNEXPECTED_TOKEN, ERR_SRC, Replace(ERR_UNEXPECTED_TOKEN_DESC, "%", p)
                End Select
            Loop
        Else
            p = p + 1
        End If
    Case "["
        ReDim v(0 To 0)
        skipWS t, p, 1
        If Mid(t, p, 1) = "]" Then
            'Empty array
            Erase v
        Else
            Do
                Assign v(UBound(v)), getValue(t, p)
                skipWS t, p
                Select Case Mid(t, p, 1)
                Case ","
                    ReDim Preserve v(0 To UBound(v) + 1)
                    skipWS t, p, 1
                Case "]" 'Last element
                    p = p + 1
                    Exit Do
                Case Else
                    Err.Raise ERR_UNEXPECTED_TOKEN, ERR_SRC, Replace(ERR_UNEXPECTED_TOKEN_DESC, "%", p)
                End Select
            Loop
        End If
    Case """"
        v = vbNullString
        Do
            p = p + 1
            k = Mid(t, p, 1)
            Select Case k
            Case """"
                p = p + 1
                Exit Do
            Case "\"
                p = p + 1
                Select Case Mid(t, p, 1)
                Case """": v = v & """"
                Case "/": v = v & "/"
                Case "\": v = v & "\"
                Case "b": v = v & vbBack
                Case "f": v = v & vbFormFeed
                Case "n": v = v & vbLf
                Case "r": v = v & vbCr
                Case "t": v = v & vbTab
                Case "u"
                    If p + 4 > Len(t) Then Err.Raise ERR_TRUNCATED, ERR_SRC, Replace(ERR_TRUNCATED_DESC, "%", p)
                    On Error GoTo RaiseInvalidEscape
                    v = v & ChrW("&H" & Mid(t, p + 1, 4))
                    On Error GoTo 0
                    p = p + 3
                Case Else
                    Err.Raise ERR_INVALID_ESCAPE, ERR_SRC, Replace(ERR_INVALID_ESCAPE_DESC, "%", p)
                End Select
            Case ""
                Err.Raise ERR_TRUNCATED, ERR_SRC, Replace(ERR_TRUNCATED_DESC, "%", p)
            Case Else
                v = v + k
            End Select
        Loop
    Case ""
        Err.Raise ERR_TRUNCATED, ERR_SRC, Replace(ERR_TRUNCATED_DESC, "%", p)
    Case Else
        'null, true, false or number
        'Get token:
        i = p
        Do
            Select Case Mid(t, p, 1)
            Case " ", vbTab, vbLf, vbCr, "", "[", "]", "{", "}", ",", ":": Exit Do
            End Select
            p = p + 1
        Loop
        k = Mid(t, i, p - i)
        Select Case k
        Case "null": v = Null
        Case "true": v = True
        Case "false": v = False
        Case Else
            If IsNumeric(Replace(k, ".", Mid(1.3, 2, 1))) And (IsNumeric(Left(k, 1)) Or Left(k, 1) = "-") Then 'Exclude &H VB formats
                v = Val(k)
            Else
                Err.Raise ERR_INVALID_VALUE, ERR_SRC, Replace(ERR_INVALID_VALUE_DESC, "%", p)
            End If
        End Select
    End Select
    If IsObject(v) Then
        Set getValue = v
    Else
        getValue = v
    End If
    Exit Function
RaiseInvalidEscape:
    Err.Raise ERR_INVALID_ESCAPE, ERR_SRC, Replace(ERR_INVALID_ESCAPE_DESC, "%", p)
End Function

'Skip white spaces helper function
Private Sub skipWS(t As String, p As Long, Optional ByVal offset As Long)
    p = p + offset
    Do While p <= Len(t)
        Select Case Mid(t, p, 1)
        Case " ", vbTab, vbLf, vbCr
        Case Else
            Exit Do
        End Select
        p = p + 1
    Loop
End Sub
