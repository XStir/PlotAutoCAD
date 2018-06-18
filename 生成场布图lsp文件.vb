Public TextHight, HangJuBL
Sub 场布图生成()

'初始化检查，如EXCEL输入有误则给出提示并停止程序
If ExcelInputBug Then
MsgBox "数据输入有误"
Exit Sub
End If

'新建lsp代码文件
Dim FileName As String
FileName = CStr(Cells(2, 13)) & ".lsp"
FileName = Replace(FileName, "/", "")
FileName = Replace(FileName, ":", "")
Open CStr(Cells(4, 13)) & CStr(ActiveSheet.Name) & FileName For Output As #1

'生成初始代码，包括：要求用户输入起始点、编制信息
Print #1, "(setq PlotPoint (getpoint ""选择插入点""))"
If TextHight < 1 Then TextHight = CStr("0" & Cells(5, 13))
If HangJuBL < 1 Then HangJuBL = CStr("0" & Cells(8, 13))

'循环开始，逐行判断是否要放置(E列)。是的话就向lsp文件中写入绘图代码并把清单中已放置(D列)设为Y，否则不做改变直接跳过。
Dim n, geshu As Integer
Dim ObjType, Objname As String
Dim L1, L2 As Single
For n = 2 To ActiveSheet.UsedRange.Rows.Count
    If Cells(n, 7) = 0 Then GoTo NXT
    If Cells(n, 5) = "Y" Then
        Objname = Cells(n, 3)
        ObjType = Cells(n, 6)
        L1 = Cells(n, 7)
        L2 = Cells(n, 8)
        geshu = 1
        If Cells(n, 9) > geshu Then geshu = Cells(n, 9)
        Cells(n, 4) = "Y"
        
        For m = 1 To geshu
            Select Case ObjType
                Case "Rectangle"
                Call LispInputRec(n, L1, L2, Objname)
                Case "Line"
                Call LispInputLine(n, L1, Objname)
                Case "Circle"
                Call LispInputCircle(n, L1, Objname)
                Case "Undefine"
                Cells(n, 4) = "N"
            End Select
        Next m
    End If
NXT: Next n

'收尾工作
Close #1
End Sub

'生成矩形的主程序
Sub LispInputRec(n, L1, L2, Objname)
'画框框
Print #1, "(setq offsetH " & L1 & " offsetV " & L2 & ")"
Print #1, "(setq Point2 (list (+ (car PlotPoint) offsetH) (+ (car (cdr PlotPoint)) offsetV)))"
Print #1, "(command ""rectang"" PlotPoint Point2)"

'写字
Dim text, quote As String
text = "\\pxqc;" & Objname & "\\P" & L1 & "x" & L2
quote = """"
text = quote & text & quote
Print #1, "(setq Point3 (list (+ (car PlotPoint) offsetH) (+ (car (cdr PlotPoint)) (* 0.5 offsetV))))"
Print #1, "(command ""mtext"" PlotPoint Point3 " & text & " """")"
Print #1, "(setq shux (entget (entlast)))(setq shux (subst (cons 40 " & TextHight & ") (assoc 40 shux) shux))(setq shux (subst (cons 44 " & HangJuBL & ") (assoc 44 shux) shux))(setq shux (subst (cons 71 4) (assoc 71 shux) shux))(entmod shux)"

'移动绘图基点
Print #1, "(setq PlotPoint (list (+ (car PlotPoint) offsetH) (car (cdr PlotPoint))))"
End Sub

'生成线的主程序
Sub LispInputLine(n, L1, Objname)
MsgBox "画线的程序还没写"
End Sub

'生成圆的主程序
Sub LispInputCircle(n, L1, Objname)
MsgBox "画圆的程序还没写"
End Sub

Function ExcelInputBug() As Boolean
ExcelInputBug = False
TextHight = Cells(5, 13)
HangJuBL = Cells(8, 13)
If TextHight * HangJuBL = 0 Then ExcelInputBug = True
End Function
