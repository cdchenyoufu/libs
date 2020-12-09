Attribute VB_Name = "modSupport"
Option Explicit
'dz: misc support functions from my library

Public Enum hexOutFormats
    hoDump
    hoSpaced
    hoHexOnly
End Enum

Function c2a(c As Collection) As String()
    Dim t() As String, f
    For Each f In c
        push t, f
    Next
    c2a = t
End Function

Function WebFileNameFromPath(fullpath)
    Dim tmp() As String
    If InStr(fullpath, "/") > 0 Then
        tmp = Split(fullpath, "/")
        WebFileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function c2s(c As Collection, Optional delimiter = vbCrLf) As String
    c2s = Join(c2a(c), delimiter)
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo Init
    Dim X
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If IsObject(value) Then
        Set ary(X + 1) = value
    Else
        ary(X + 1) = value
    End If
    
    Exit Sub
Init:
    ReDim ary(0)
    If IsObject(value) Then
        Set ary(0) = value
    Else
        ary(0) = value
    End If
End Sub

Function FileExists(path) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function GetParentFolder(path, Optional levelUp = 1)
    Dim tmp() As String
    Dim my_path
    Dim ub As String, i As Long
    
    On Error GoTo hell
    If Len(path) = 0 Then Exit Function
    If levelUp < 1 Then levelUp = 1
    
    my_path = path
    While Len(my_path) > 0 And Right(my_path, 1) = "\"
        my_path = Mid(my_path, 1, Len(my_path) - 1)
    Wend
    
    tmp = Split(my_path, "\")
    If levelUp > UBound(tmp) Then levelUp = UBound(tmp)
    
    For i = 0 To levelUp - 1
        If InStr(tmp(UBound(tmp) - i), ":") < 1 Then tmp(UBound(tmp) - i) = Empty
    Next
    
    my_path = Join(tmp, "\")
    While Len(my_path) > 0 And Right(my_path, 1) = "\"
        my_path = Mid(my_path, 1, Len(my_path) - 1)
    Wend
        
    GetParentFolder = my_path
    Exit Function
    
hell:
    GetParentFolder = Empty
    
End Function

Function DeleteFile(fPath)
 On Error GoTo hadErr
    
    Dim attributes As VbFileAttribute

    attributes = GetAttr(fPath)
    If (attributes And vbReadOnly) Then
        attributes = attributes - vbReadOnly
        SetAttr fPath, attributes
    End If
    
    Kill fPath
    DeleteFile = True
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim X
    If IsObject(UBound(ary)) Then AryIsEmpty = False
    'x = UBound(ary)
  Exit Function
oops: AryIsEmpty = True
End Function

Function HexDump(bAryOrStrData, Optional ByVal Length As Long = -1, Optional ByVal startAt As Long = 1, Optional hexFormat As hexOutFormats = hoDump) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, X
    Dim hexOnly As Long
    
    offset = 0
    If hexFormat <> hoDump Then hexOnly = 1
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        If AryIsEmpty(bAryOrStrData) Then Exit Function
        ary() = bAryOrStrData
    Else
        If Len(CStr(bAryOrStrData)) = 0 Then Exit Function
        ary = StrConv(CStr(bAryOrStrData), vbFromUnicode, LANG_US)
    End If
    
    If startAt < 1 Then startAt = 1
    If Length < 1 Then Length = -1
    
    While startAt Mod 16 <> 0
        startAt = startAt - 1
    Wend
    
    startAt = startAt + 1
    
    chars = "   "
    For i = startAt To UBound(ary) + 1
        tt = Hex(ary(i - 1))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        X = ary(i - 1)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((X > 32 And X < 127), Chr(X), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
        If Length <> -1 Then
            Length = Length - 1
            If Length = 0 Then Exit For
        End If
    Next
    
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        If hexFormat = hoHexOnly Then HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

