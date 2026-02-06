<%
' =================================================================
' Professional Benchmark Class - Final Article Edition
' Optimized with & _ line continuation and expanded VBScript syntax.
' =================================================================
Class Benchmark_Class
    Private p_Name
    Private p_Mode
    Private p_ModeValue
    Private p_Lock
    Private p_HaltOnError
    Private p_LogTo
    Private p_ErrorTo
    Private p_DataTo
    Private p_Rounds
    Private p_InitGlobal
    Private p_InitBlock
    Private p_CodeBlocks
    Private p_Labels
    Private p_FirstRun
    Private fso
    Private p_LogBuffer
    Private p_ErrorBuffer
    Private p_DataBuffer

    Private Sub Class_Initialize()
        On Error Resume Next
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        p_CodeBlocks = Array()
        p_Labels = Array()
        p_FirstRun = True
        p_Rounds = 1
        p_LogBuffer = ""
        p_ErrorBuffer = ""
        p_DataBuffer = ""
        p_Name = "default"
        p_Mode = "COUNT"
        p_ModeValue = 100000
        p_Lock = False
        p_HaltOnError = False
        p_LogTo = "SCREEN"
        p_ErrorTo = "SCREEN"
        p_DataTo = "FILE"
    End Sub

    Private Sub Class_Terminate()
        Set fso = Nothing
    End Sub

    ' --- Properties ---
    Public Property Let Name(v)
        p_Name = v
    End Property

    Public Property Let Rounds(v)
        p_Rounds = CLng(v)
    End Property

    Public Property Let Lock(v)
        If IsBoolean(v) Then
            p_Lock = v
        End If
    End Property

    Public Property Let HaltOnError(v)
        If IsBoolean(v) Then
            p_HaltOnError = v
        End If
    End Property

    Public Property Let LogTo(v)
        p_LogTo = ValidateOutput(v, p_LogTo)
    End Property

    Public Property Let ErrorTo(v)
        p_ErrorTo = ValidateOutput(v, p_ErrorTo)
    End Property

    Public Property Let DataTo(v)
        p_DataTo = ValidateOutput(v, p_DataTo)
    End Property

    Public Property Let InitGlobal(v)
        p_InitGlobal = LoadCode(v)
    End Property

    Public Property Let InitBlock(v)
        p_InitBlock = LoadCode(v)
    End Property

    ' --- Public Getters ---
    Public Function getLog()
        getLog = p_LogBuffer
    End Function

    Public Function getError()
        getError = p_ErrorBuffer
    End Function

    Public Function getData()
        getData = p_DataBuffer
    End Function

    ' --- Public Methods ---
    Public Sub SetMode(m, v)
        Dim uM
        uM = UCase(m)
        If (uM = "COUNT" Or uM = "TIME") And IsNumeric(v) Then
            p_Mode = uM
            p_ModeValue = CLng(v)
        Else
            HandleErr "Invalid Mode/Value."
        End If
    End Sub

    Public Sub AddCodeBlock(lbl, code)
        Dim c
        c = LoadCode(code)
        If c <> "" Then
            Dim i
            On Error Resume Next
            i = UBound(p_CodeBlocks)
            If Err.Number <> 0 Then
                i = -1
            End If
            On Error GoTo 0
            i = i + 1
            ReDim Preserve p_CodeBlocks(i)
            ReDim Preserve p_Labels(i)
            p_CodeBlocks(i) = c
            p_Labels(i) = p_IIf(lbl = "" And Left(code, 5) = "File:", Mid(code, 6), lbl)
            Log "Block Added: " & p_Labels(i)
        End If
    End Sub

    Public Sub Run()
        Dim bm_r, bm_i, bm_start, bm_iterations, bm_duration, bm_best, bm_res, bm_err
        If p_FirstRun Then
            InternalWarmup()
            p_FirstRun = False
        End If
        If p_Lock Then
            Application.Lock
        End If
        Err.Clear
        On Error Resume Next 
        ExecuteGlobal p_InitGlobal
        If Err.Number <> 0 Then
            HandleErr "InitGlobal: " & Err.Description
            If p_Lock Then
                Application.UnLock
            End If
            Exit Sub
        End If
        bm_res = p_Name & Chr(31) & p_Rounds & Chr(29)
        For bm_i = 0 To UBound(p_CodeBlocks)
            Log "Testing: " & p_Labels(bm_i) & " (" & p_Rounds & " rounds)..."
            bm_best = 999999
            bm_err = False
            For bm_r = 1 To p_Rounds
                bm_iterations = 0
                If p_Mode = "COUNT" Then
                    bm_start = Timer
                    For bm_iterations = 1 To p_ModeValue
                        ExecuteGlobal p_InitBlock
                        ExecuteGlobal p_CodeBlocks(bm_i)
                        If bm_r = 1 And bm_iterations = 1 And Err.Number <> 0 Then
                            bm_err = True
                            Exit For
                        End If
                    Next
                    bm_duration = Timer - bm_start
                Else
                    bm_start = Timer
                    Do While (Timer - bm_start) < p_ModeValue
                        ExecuteGlobal p_InitBlock
                        ExecuteGlobal p_CodeBlocks(bm_i)
                        bm_iterations = bm_iterations + 1
                        If bm_r = 1 And bm_iterations = 1 And Err.Number <> 0 Then
                            bm_err = True
                            Exit Do
                        End If
                    Loop
                    bm_duration = Timer - bm_start
                End If
                If bm_duration < bm_best Then
                    bm_best = bm_duration
                End If
                Log "Finished Round " & bm_r & ": " & p_Labels(bm_i)
                If bm_err Then
                    Exit For
                End If
            Next
            If Not bm_err Then
                bm_res = bm_res & p_Labels(bm_i) & Chr(31) & bm_iterations & Chr(31) & Round(bm_best, 4) & Chr(30)
            Else
                HandleErr "Runtime Error in [" & p_Labels(bm_i) & "]"
            End If
            If p_LogTo = "SCREEN" Then
                Response.Flush 
            End If
        Next
        WriteData bm_res
        If p_Lock Then
            Application.UnLock
        End If
    End Sub

    ' --- Optimized UI Helpers ---
    Public Function getHTMLTable()
        If p_DataBuffer = "" Then
            getHTMLTable = "<p class=""benchmark-no-data"">No data available.</p>"
            Exit Function
        End If
        Dim bm_rows, bm_head, bm_cells, bm_idx, bm_html, bm_rec, bm_meta, bm_list, bm_winTime, bm_curTime, bm_rank, bm_parts, bm_factor
        Set bm_list = CreateObject("System.Collections.ArrayList")
        bm_head = Split(p_DataBuffer, Chr(29))
        bm_meta = Split(bm_head(0), Chr(31))
        bm_rows = Split(bm_head(1), Chr(30))
        For bm_idx = 0 To UBound(bm_rows)
            bm_rec = Trim(bm_rows(bm_idx))
            If bm_rec <> "" Then
                bm_cells = Split(bm_rec, Chr(31))
                bm_list.Add Right("000000" & FormatNumber(bm_cells(2), 4, -1, 0, 0), 12) & "|" & bm_cells(0) & "|" & bm_cells(1)
            End If
        Next
        bm_list.Sort()
        If bm_list.Count > 0 Then
            bm_winTime = CDbl(Split(bm_list(0), "|")(0))
        End If
        bm_html = "<div class=""benchmark-result-container"">" & _
                  "<table class=""benchmark-table"">" & _
                  "<thead>" & _
                  "<tr><th colspan='4'>Benchmark: " & Server.HTMLEncode(bm_meta(0)) & " (" & bm_meta(1) & " rounds)</th></tr>" & _
                  "<tr><th>Rank / Label</th><th>Iterations</th><th>Time</th><th>Factor</th></tr>" & _
                  "</thead>" & _
                  "<tbody>"
        bm_rank = 0
        For Each bm_rec In bm_list
            bm_parts = Split(bm_rec, "|")
            bm_curTime = CDbl(bm_parts(0))
            bm_rank = bm_rank + 1
            If bm_winTime > 0 Then
                bm_factor = Round(bm_curTime / bm_winTime, 2)
            Else
                bm_factor = 1
            End If
            bm_html = bm_html & "<tr>" & _
                      "<td>#" & bm_rank & " <b>" & Server.HTMLEncode(bm_parts(1)) & "</b></td>" & _
                      "<td>" & FormatNumber(bm_parts(2), 0) & "</td>" & _
                      "<td>" & bm_curTime & "s</td>"
            If bm_rank = 1 Then
                bm_html = bm_html & "<td class=""benchmark-factor""><b>1.0x (Fastest)</b></td>"
            Else
                bm_html = bm_html & "<td class=""benchmark-factor"">" & bm_factor & "x slower</td>"
            End If
            bm_html = bm_html & "</tr>"
        Next
        getHTMLTable = bm_html & "</tbody></table></div>"
        Set bm_list = Nothing
    End Function

    Public Function Help()
        Help = "<div class=""benchmark-help"">" & _
               "<h3 class=""benchmark-title"">Benchmark API Help</h3>" & _
               "<p><b>Modes:</b> SCREEN, FILE, STR, CASHE, CASHE_FILE, NONE</p>" & _
               "<ul class=""benchmark-props"">" & _
               "<li>Properties: .Name, .Rounds, .Lock, .LogTo, .DataTo</li>" & _
               "</ul>" & _
               "<pre class=""benchmark-methods"">" & _
               ".SetMode(mode, val)  ' mode: ""COUNT"" or ""TIME""" & vbCrLf & _
               ".AddCodeBlock(lbl, c) ' c: snippet or ""File:filename""" & vbCrLf & _
               ".Run()               ' Executes benchmark" & vbCrLf & _
               ".getHTMLTable()       ' Returns results HTML" & _
               "</pre></div>"
    End Function

    Public Function Vars()
        Vars = "<div class=""benchmark-vars"">" & _
               "<span class=""benchmark-status""><b>Benchmark State:</b> " & _
               "Name: " & p_Name & " | Rounds: " & p_Rounds & _
               " | Mode: " & p_Mode & " (" & p_ModeValue & ") | " & _
               "Lock: " & p_Lock & _
               "</span></div>"
    End Function

    ' --- Private Internals ---
    Private Sub Log(msg)
        If p_LogTo = "SCREEN" Then
            Response.Write "[LOG] " & msg & "<br>"
        End If
        If p_LogTo = "STR" Then
            p_LogBuffer = p_LogBuffer & msg & vbCrLf
        End If
    End Sub

    Private Sub HandleErr(msg)
        Dim fE
        fE = "<b style='color:red;'>[ERR] " & msg & "</b><br>"
        If p_ErrorTo = "SCREEN" Then
            Response.Write fE
        End If
        If p_ErrorTo = "STR" Then
            p_ErrorBuffer = p_ErrorBuffer & msg & vbCrLf
        End If
        If p_HaltOnError Then
            If p_Lock Then
                Application.UnLock
            End If
            Response.End
        End If
    End Sub

    Private Sub WriteData(data)
        p_DataBuffer = data 
        If p_DataTo = "FILE" Or p_DataTo = "CASHE_FILE" Then
            On Error Resume Next
            fso.CreateTextFile(Server.MapPath("examples/" & p_Name & ".dat"), True).Write data
        End If
        If (p_DataTo = "CASHE" Or p_DataTo = "CASHE_FILE") Then
            Dim l
            l = Not p_Lock
            If l Then
                Application.Lock
            End If
            Application("Bench_Live_" & p_Name) = data
            If l Then
                Application.UnLock
            End If
        End If
        If p_DataTo = "SCREEN" Then
            Response.Write getHTMLTable()
        End If
    End Sub

    Private Sub InternalWarmup()
        Log "System: Warming Up..."
        Dim s
        s = Timer
        Do While (Timer - s) < 0.5
        Loop
    End Sub

    Private Function LoadCode(v)
        If Left(v, 5) = "File:" Then
            Dim p
            p = Server.MapPath("examples/" & Trim(Mid(v, 6)))
            If fso.FileExists(p) Then
                LoadCode = fso.OpenTextFile(p, 1).ReadAll
            Else
                HandleErr "File missing: " & v
            End If
        Else
            LoadCode = v
        End If
    End Function

    Private Function ValidateOutput(v, fb)
        Dim o
        o = UCase(v)
        Select Case o
            Case "FILE", "SCREEN", "STR", "CASHE", "CASHE_FILE", "NONE"
                ValidateOutput = o
            Case Else
                ValidateOutput = fb
        End Select
    End Function

    Private Function IsBoolean(v)
        IsBoolean = (VarType(v) = 11)
    End Function

    Private Function p_IIf(e, t, f)
        If e Then
            p_IIf = t
        Else
            p_IIf = f
        End If
    End Function
End Class
%>
