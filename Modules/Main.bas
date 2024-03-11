Attribute VB_Name = "Main"
Option Explicit

Dim AppDict As Object
Dim ObjPropDict As Object

Public Function YQ(Base As String, Query As String, ParamArray Params() As Variant) As Variant
    On Error GoTo HandleError
    Dim App As Object
    Dim Parameters() As Variant
    Dim Ret As Object, RetRow As Object, Column As Object, RowNum As Long, ColNum As Long
    Dim ResultArray() As Variant
    Dim I As Long
    
    Set App = GetAppByRef(Base)
    If LBound(Params) <= UBound(Params) Then
        ReDim Parameters(UBound(Params))
        For I = LBound(Params) To UBound(Params)
            If TypeName(Params(I)) = "Range" Then
                Parameters(I) = Params(I).Value
            Else
                Parameters(I) = Params(I)
            End If
        Next I
        Set Ret = App.YQ_OLEAutomationClient.RunQuery(Query, Parameters)
    Else
        Set Ret = App.YQ_OLEAutomationClient.RunQuery(Query)
    End If
    If Ret.IsArray Then
        RowNum = 0
        ReDim ResultArray(Ret.RowCount - 1, Ret.ColumnCount - 1)
        For Each RetRow In Ret.Value
            ColNum = 0
            For Each Column In RetRow
                ResultArray(RowNum, ColNum) = Column.Value
                ColNum = ColNum + 1
            Next
            RowNum = RowNum + 1
        Next
        YQ = ResultArray
    ElseIf Ret.RowCount > 0 Then
        YQ = Ret.Value
    Else
        YQ = CVErr(xlErrNA)
    End If
    Exit Function
HandleError:
    Debug.Print "Main.YQ", "Èñêëþ÷åíèå", Err.Number, Err.Description
    YQ = CVErr(xlErrValue)
End Function

Public Function ÏÐÅÄÑÒÀÂËÅÍÈÅÑÑÛËÊÈ(Rng As Variant) As Variant
    On Error GoTo HandleError
    Dim Result As Variant
    Dim App As Object, Ref As String, DictKey As String
    
    If TypeName(Rng) = "Range" Then
        Ref = FindRefInRange(Rng)
    Else
        Ref = Rng
    End If
    If Ref = "" Then
        Result = CVErr(xlErrValue)
    Else
        If ObjPropDict Is Nothing Then
            Set ObjPropDict = CreateObject("Scripting.Dictionary")
        End If
        DictKey = Ref + "_View"
        If ObjPropDict.exists(DictKey) Then
            Result = ObjPropDict(DictKey)
        Else
            Set App = GetAppByRef(Ref)
            Result = App.YQ_OLEAutomationClient.GetURLPresentation(Ref)
            ObjPropDict.Add DictKey, Result
        End If
    End If
    ÏÐÅÄÑÒÀÂËÅÍÈÅÑÑÛËÊÈ = Result
    Exit Function
HandleError:
    Debug.Print "Main.ÏÐÅÄÑÒÀÂËÅÍÈÅÑÑÛËÊÈ", "Èñêëþ÷åíèå", Err.Number, Err.Description
    ÏÐÅÄÑÒÀÂËÅÍÈÅÑÑÛËÊÈ = CVErr(xlErrValue)
End Function

Public Function ÐÅÊÂÈÇÈÒÑÑÛËÊÈ(Rng As Range, PropertyName As String) As Variant
    On Error GoTo HandleError
    Dim Result As Variant
    Dim App As Object, DictKey As String
    Dim Ref As String
    
    If PropertyName = "" Then
        Result = CVErr(xlErrValue)
        Exit Function
    End If
    If TypeName(Rng) = "Range" Then
        Ref = FindRefInRange(Rng)
    Else
        Ref = Rng
    End If
    If Ref = "" Then
        Result = CVErr(xlErrValue)
    Else
        If ObjPropDict Is Nothing Then
            Set ObjPropDict = CreateObject("Scripting.Dictionary")
        End If
        DictKey = Ref + "." + PropertyName
        If ObjPropDict.exists(DictKey) Then
            Result = ObjPropDict(DictKey)
        Else
            Set App = GetAppByRef(Ref)
            Result = App.YQ_OLEAutomationClient.GetURLProperty(Ref, PropertyName)
            ObjPropDict.Add DictKey, Result
        End If
    End If
    ÐÅÊÂÈÇÈÒÑÑÛËÊÈ = Result
    Exit Function
HandleError:
    Debug.Print "Main.ÐÅÊÂÈÇÈÒÑÑÛËÊÈ", "Èñêëþ÷åíèå", Err.Number, Err.Description
    ÐÅÊÂÈÇÈÒÑÑÛËÊÈ = CVErr(xlErrValue)
End Function

Private Function FindRefInRange(Rng As Variant) As String
    Dim Result As String
    
    Result = FindRefInStr(Rng.Value2)
    If Result = "" Then
        Result = FindRefInStr(Rng.Formula2)
    End If
    FindRefInRange = Result
End Function

Private Function FindRefInStr(Str As String) As String
    Dim Result As String
    Dim Pos As Long, Pos1 As Long
    
    Pos = InStr(Str, "e1c")
    If Pos > 0 Then
        Pos1 = InStr(Pos, Str, """")
        If Pos1 > 0 Then
            Result = Mid(Str, Pos, Pos1 - Pos)
        Else
            Result = Mid(Str, Pos)
        End If
    End If
    FindRefInStr = Result
End Function

Private Function GetAppByRef(Ref As String) As Object
    Dim App As Object
    Dim ConStr As String
    
    ConStr = GetConStrByRef(Ref)
    Set App = GetApp(ConStr)
    Set GetAppByRef = App
End Function

Private Function GetConStrByRef(Ref As String) As String
    Dim Base As String, Result As String
    
    Result = ""
    If Ref <> "" Then
        Base = GetBaseByRef(Ref)
        Result = GetConStrByBase(Base)
    End If
    GetConStrByRef = Result
End Function

Private Function GetBaseByRef(Ref As String) As String
    Dim Pos As Long, Pos1 As Long
    Dim Result As String
    
    Result = ""
    If Ref <> "" Then
        Pos = InStr(Ref, "server/")
        If Pos > 0 Then
            Pos1 = InStr(Pos, Ref, "#")
            If Pos1 > 0 Then
                Result = Mid(Ref, Pos, Pos1 - Pos)
            Else
                Result = Mid(Ref, Pos)
            End If
        Else
            Pos = InStr(Ref, "filev/")
            If Pos > 0 Then
                Pos1 = InStr(Pos, Ref, "#")
                If Pos1 > 0 Then
                    Result = Mid(Ref, Pos, Pos1 - Pos)
                Else
                    Result = Mid(Ref, Pos)
                End If
            End If
        End If
    End If
    GetBaseByRef = Result
End Function

Private Function GetConStrByBase(Base As String) As String
    Dim Pos As Long, Pos1 As Long
    Dim Server As String, Ib As String, File As String
    Dim Result As String
    
    Result = ""
    If Base <> "" Then
        If Left(Base, 7) = "server/" Then
            Pos = 8
            Pos1 = InStr(Pos, Base, "/")
            If Pos1 > 0 Then
                Server = Mid(Base, Pos, Pos1 - Pos)
                Ib = Mid(Base, Pos1 + 1)
                Result = "Srvr=""" + Server + """;Ref=""" + Ib + """;"
            End If
        ElseIf Left(Base, 6) = "filev/" Then
            Pos = 7
            File = Mid(Base, Pos)
            File = Replace(File, "/", "\")
            File = Replace(File, "\", ":\", , 1)
            Result = "File=""" + File + """;"
        End If
    End If
    GetConStrByBase = Result
End Function

Private Function GetApp(ConStr As String) As Object
    On Error GoTo HandleError
    Dim App As Object, Ret As Boolean
    
    If AppDict Is Nothing Then
        Set AppDict = CreateObject("Scripting.Dictionary")
    End If
    If ConStr = "" Then
        Exit Function
    End If
    If AppDict.exists(ConStr) Then
        Set App = AppDict(ConStr)
    End If
    If App Is Nothing Then
        Set App = CreateObject("V83C.Application")
        Ret = App.Connect(ConStr)
        Debug.Print "Main.GetApp", "Ïîäêëþ÷åíèå", ConStr, Ret
        AppDict.Add ConStr, App
    End If
    Set GetApp = App
    Exit Function
HandleError:
    Debug.Print "Main.GetApp", "Èñêëþ÷åíèå", Err.Number, Err.Description
End Function
