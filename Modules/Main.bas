Attribute VB_Name = "Main"
Option Explicit

Dim AppDict As Object
Dim ObjPropDict As Object

Public Function YQ(Base As String, QueryText As String, ParamArray Params() As Variant) As Variant
Attribute YQ.VB_Description = "Возвращает результат выполнения запроса на языке 1С"
Attribute YQ.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error GoTo HandleError
    Dim App As Object
    Dim Parameters() As Variant
    Dim Ret As Object
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
        Set Ret = App.YQ_OLEAutomationClient.RunQuery(QueryText, Parameters)
    Else
        Set Ret = App.YQ_OLEAutomationClient.RunQuery(QueryText)
    End If
    If Ret.IsArray Then
        YQ = Ret.Value
    ElseIf Ret.RowCount > 0 Then
        YQ = Ret.Value
    Else
        YQ = CVErr(xlErrNA)
    End If
    'Debug.Print "Main.YQ", "Время выполнения запроса, с", Ret.Duration
    Exit Function
HandleError:
    Debug.Print "Main.YQ", "Исключение", Err.Number, Err.Description
    YQ = CVErr(xlErrValue)
End Function

Public Function REFP(Reference As Variant) As Variant
Attribute REFP.VB_Description = "Получает представление ссылки"
Attribute REFP.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error GoTo HandleError
    Dim Result As Variant
    Dim App As Object, Ref As String, DictKey As String
    
    If TypeName(Reference) = "Range" Then
        Ref = FindRefInRange(Reference)
    Else
        Ref = Reference
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
    REFP = Result
    Exit Function
HandleError:
    Debug.Print "Main.REFP", "Исключение", Err.Number, Err.Description
    REFP = CVErr(xlErrValue)
End Function

Public Function REFA(Reference As Variant, AttributeName As String) As Variant
Attribute REFA.VB_Description = "Получает значение реквизита ссылки"
Attribute REFA.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error GoTo HandleError
    Dim Result As Variant
    Dim App As Object, DictKey As String
    Dim Ref As String
    
    If AttributeName = "" Then
        Result = CVErr(xlErrValue)
        Exit Function
    End If
    If TypeName(Reference) = "Range" Then
        Ref = FindRefInRange(Reference)
    Else
        Ref = Reference
    End If
    If Ref = "" Then
        Result = CVErr(xlErrValue)
    Else
        If ObjPropDict Is Nothing Then
            Set ObjPropDict = CreateObject("Scripting.Dictionary")
        End If
        DictKey = Ref + "." + AttributeName
        If ObjPropDict.exists(DictKey) Then
            Result = ObjPropDict(DictKey)
        Else
            Set App = GetAppByRef(Ref)
            Result = App.YQ_OLEAutomationClient.GetURLAttribute(Ref, AttributeName)
            ObjPropDict.Add DictKey, Result
        End If
    End If
    REFA = Result
    Exit Function
HandleError:
    Debug.Print "Main.REFA", "Исключение", Err.Number, Err.Description
    REFA = CVErr(xlErrValue)
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
        Debug.Print "Main.GetApp", "Подключение", ConStr, Ret
        AppDict.Add ConStr, App
    End If
    Set GetApp = App
    Exit Function
HandleError:
    Debug.Print "Main.GetApp", "Исключение", Err.Number, Err.Description
End Function

Public Sub RegisterYQFunctions()
    Dim YQArgs(1 To 3) As Variant, REFPArgs(1 To 1) As Variant, REFAArgs(1 To 2) As Variant
    YQArgs(1) = "Навигационная ссылка информационной базы 1С, к которой выполняется запрос"
    YQArgs(2) = "Текст запроса на языке 1С"
    YQArgs(3) = "Произвольное число параметров запроса. Параметры задаются парами имя;значение"
    Application.MacroOptions Macro:="YQ", Description:="Возвращает результат выполнения запроса на языке 1С", ArgumentDescriptions:=YQArgs
    REFPArgs(1) = "Внешняя навигационная ссылка объекта, для которого требуется получить представление"
    Application.MacroOptions Macro:="REFP", Description:="Получает представление ссылки", ArgumentDescriptions:=REFPArgs
    REFAArgs(1) = "Внешняя навигационная ссылка объекта, значение реквизита которого требуется получить"
    REFAArgs(2) = "Имя реквизита объекта, значение которого требуется получить"
    Application.MacroOptions Macro:="REFA", Description:="Получает значение реквизита ссылки", ArgumentDescriptions:=REFAArgs
End Sub
