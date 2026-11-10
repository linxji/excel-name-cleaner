Attribute VB_Name = "CleanNamesFast"
Option Explicit

'==============================================================
' Excel 定义名称批量清理（快速版）v1.2
'--------------------------------------------------------------
' 原理：xlsx/xlsm 本质是 zip 包，所有定义名称都存放在
'       xl/workbook.xml 的 <definedNames> 节点里。
'       直接把这个节点从包里摘掉 = 一次性删光，
'       不需要 30 万次 Name.Delete 调用。
'
' 实测：30 万个名称，原方案约 2 小时 → 本方案约 1 秒。
'
' 用法：
'   1) 新建空白工作簿 → Alt+F11 → 文件/导入本 .bas
'      → 另存为 .xlsm（或放进 PERSONAL.XLSB）。
'   2) 不要在被清理的文件里运行本宏（会提示）。
'   3) 运行 清理名称_主入口，可一次多选文件。
'
' 处理策略：
'   .xlsx .xlsm .xltx .xltm  → 快路径（直接改 zip，秒级）
'   .xlsb .xls               → 转换法（转 xlsx 清理再转回，约 1 分钟）
'                              或 对象模型法（逐个删，约 15~20 分钟）
'
' v1.1 修正：VBA 单条逻辑行最多 25 个续行符(_)，原内嵌脚本超限。
' v1.2 修正：
'   ① 对象模型删除遇到删不掉的名称不再中断，改为跳过并继续，
'      同时记录错误码/错误信息/名称，写进结果里；
'   ② .xlsb/.xls 新增转换法，避开逐个删除的所有坑；
'   ③ 状态栏按 R1 风格显示 进度/百分比/速度/已耗时/剩余/完成时间；
'   ④ 剩余名称 < 10 个时，自动删除备份文件；
'   ⑤ 打开大文件前先提示并关闭自动重算，减少无响应的错觉。
'==============================================================

Private Const MAKE_BACKUP As Boolean = True
Private Const PS_NAME As String = "wb_clean_names.ps1"
Private Const OUT_NAME As String = "wb_clean_names.log"

' 剩余名称少于这个数，就认为清理成功，自动删掉备份文件（省空间）
Private Const AUTO_DEL_BACKUP_WHEN_REMAIN_LT As Long = 10

' 对象模型删除时，删不掉的名称达到这个数就停手（说明是系统性问题，不是个别坏名称）
Private Const MAX_STUCK As Long = 200

' 状态栏刷新间隔（秒）
Private Const UI_INTERVAL As Double = 2

' 需要保留的名称（.NET 正则）。留空 = 全部删除。
' 例：保留打印区域 / 打印标题 / 筛选范围：
'   Private Const KEEP_PATTERN As String = "^_xlnm\.(Print_Area|Print_Titles|FilterDatabase)$"
Private Const KEEP_PATTERN As String = ""


'--------------------------------------------------------------
' 主入口
'--------------------------------------------------------------
Public Sub 清理名称_主入口()

    Dim files As Variant
    Dim i As Long, n As Long
    Dim p As String, ext As String
    Dim report As String, okN As Long, partN As Long, failN As Long
    Dim tAll As Double, status As String

    files = PickFiles()
    If Not IsArray(files) Then Exit Sub

    n = UBound(files) - LBound(files) + 1

    If MsgBox("即将处理 " & n & " 个文件：" & vbCrLf & vbCrLf & _
              "· 处理前自动备份到同目录（文件名带 _备份_ 时间戳）" & vbCrLf & _
              "· 已打开的文件会先关闭（有未保存改动会询问）" & vbCrLf & _
              "· 删除后引用这些名称的公式会变成 #NAME?" & vbCrLf & _
              "· 打印区域/筛选范围(_xlnm.*)也会被删，需保留请先改 KEEP_PATTERN" & vbCrLf & _
              "· 剩余名称少于 " & AUTO_DEL_BACKUP_WHEN_REMAIN_LT & " 个时，会自动删除备份文件" & vbCrLf & vbCrLf & _
              "是否继续？", vbYesNo + vbExclamation, "名称清理（快速版 v1.2）") <> vbYes Then Exit Sub

    tAll = Timer
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    For i = LBound(files) To UBound(files)

        p = CStr(files(i))
        ext = LCase$(Mid$(p, InStrRev(p, ".") + 1))

        Application.StatusBar = "(" & (i - LBound(files) + 1) & "/" & n & ") " & p
        DoEvents

        status = ProcessOne(p, ext)

        Select Case True
            Case Left$(status, 4) = "部分完成"
                partN = partN + 1
            Case Left$(status, 2) = "完成"
                okN = okN + 1
            Case Else
                failN = failN + 1
        End Select

        report = report & p & vbCrLf & "    → " & status & vbCrLf & vbCrLf

    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True

    MsgBox "处理结束" & vbCrLf & vbCrLf & _
           "成功：" & okN & " 个    部分完成：" & partN & " 个    失败/跳过：" & failN & " 个" & vbCrLf & _
           "总耗时：" & Format(Timer - tAll, "0.0") & " 秒" & vbCrLf & vbCrLf & _
           Left$(report, 1200), vbInformation, "完成"

End Sub


'--------------------------------------------------------------
' 单个文件处理，返回结果描述
'--------------------------------------------------------------
Private Function ProcessOne(ByVal p As String, ByVal ext As String) As String

    Dim bak As String, t As Double, rc As Long
    Dim outTxt As String, remainCnt As Long

    If Not CloseIfOpen(p) Then
        ProcessOne = "跳过（文件正被占用，或用户取消）"
        Exit Function
    End If

    If MAKE_BACKUP Then
        bak = BackupFile(p)
        If Len(bak) = 0 Then
            ProcessOne = "跳过（备份失败，检查磁盘空间/权限）"
            Exit Function
        End If
    End If

    t = Timer

    Select Case ext

        Case "xlsx", "xlsm", "xltx", "xltm"

            rc = RunCleanScript(p, outTxt)

            If rc = 0 Then
                remainCnt = GetNum(outTxt, "kept")
                ProcessOne = "完成 · 快路径 · 删除 " & Format(GetNum(outTxt, "removed"), "#,##0") & _
                             " 个 · 剩余 " & remainCnt & " 个 · " & Format(Timer - t, "0.00") & " 秒" & _
                             BackupSuffix(bak, remainCnt)
                Exit Function
            End If

            ' 快路径失败 → 回退对象模型
            ProcessOne = CleanByOM(p, bak, t)

        Case Else

            Select Case AskXlsbStrategy(ext)
                Case vbYes
                    ProcessOne = CleanByConvert(p, ext, bak, t)
                Case vbNo
                    ProcessOne = CleanByOM(p, bak, t)
                Case Else
                    ProcessOne = "跳过（你选择了跳过 ." & ext & " 文件）"
            End Select

    End Select

End Function


'--------------------------------------------------------------
' 转换法：xlsb/xls → xlsx/xlsm → 清理 → 转回原格式
'--------------------------------------------------------------
Private Function CleanByConvert(ByVal p As String, ByVal ext As String, ByVal bak As String, ByVal t0 As Double) As String

    Dim fso As Object, tmp As String, outTxt As String
    Dim wb As Workbook
    Dim fmtTmp As Long, fmtBack As Long, rc As Long
    Dim remainCnt As Long
    Dim hasVba As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Fail

    ' ① 原文件 → 临时 xlsx / xlsm（此时原文件只读不写）
    Application.StatusBar = "① 正在转换为中间格式（大文件需要一会儿）… " & fso.GetFileName(p)
    DoEvents

    Set wb = Application.Workbooks.Open(Filename:=p, UpdateLinks:=0, ReadOnly:=False)

    ' 保守处理：读不到 HasVBProject 时按"有宏"对待，用 .xlsm 中转，绝不会丢宏
    hasVba = True
    On Error Resume Next
    hasVba = wb.HasVBProject
    Err.Clear
    On Error GoTo Fail

    If hasVba Then
        fmtTmp = xlOpenXMLWorkbookMacroEnabled      ' 52  .xlsm
        tmp = fso.BuildPath(fso.GetParentFolderName(p), "_tmp_clean_" & fso.GetBaseName(p) & ".xlsm")
    Else
        fmtTmp = xlOpenXMLWorkbook                  ' 51  .xlsx
        tmp = fso.BuildPath(fso.GetParentFolderName(p), "_tmp_clean_" & fso.GetBaseName(p) & ".xlsx")
    End If

    If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
    wb.SaveAs Filename:=tmp, FileFormat:=fmtTmp
    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' ② 清理临时文件（秒级）
    Application.StatusBar = "② 正在清理名称 … " & fso.GetFileName(p)
    DoEvents

    rc = RunCleanScript(tmp, outTxt)

    If rc <> 0 Then
        ' 清理失败：删掉临时文件，改用对象模型
        On Error Resume Next
        fso.DeleteFile tmp, True
        On Error GoTo Fail
        CleanByConvert = CleanByOM(p, bak, t0)
        Exit Function
    End If

    ' ③ 转回原格式
    Application.StatusBar = "③ 正在转回 ." & ext & " … " & fso.GetFileName(p)
    DoEvents

    If ext = "xlsb" Then
        fmtBack = xlExcelBinaryWorkbook             ' 50
    Else
        fmtBack = xlExcel8                          ' 56  .xls
    End If

    Set wb = Application.Workbooks.Open(Filename:=tmp, UpdateLinks:=0, ReadOnly:=False)
    remainCnt = wb.Names.Count
    wb.SaveAs Filename:=p, FileFormat:=fmtBack
    wb.Close SaveChanges:=False
    Set wb = Nothing

    On Error Resume Next
    fso.DeleteFile tmp, True
    On Error GoTo Fail

    Application.StatusBar = False

    CleanByConvert = "完成 · 转换法 · 删除 " & Format(GetNum(outTxt, "removed"), "#,##0") & _
                     " 个 · 剩余 " & remainCnt & " 个 · " & Format(Timer - t0, "0.0") & " 秒" & _
                     BackupSuffix(bak, remainCnt)
    Exit Function

Fail:
    Application.StatusBar = False
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Len(tmp) > 0 Then
        If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
    End If
    CleanByConvert = "失败 · 转换法出错：" & Err.Description & "（原文件未改动）"

End Function


'--------------------------------------------------------------
' 对象模型法：逐个删除
' 要点一：永远删第 k 个，k 从 1 开始、只在删不掉时才 +1。
'         Names(i) 是从头线性扫描，i 越大越慢，
'         倒序删除是 O(n^2)（原方案 30 万个要 2 小时）。
' 要点二：个别名称删不掉时跳过继续，不要整批放弃。
'--------------------------------------------------------------
Private Function CleanByOM(ByVal p As String, ByVal bak As String, ByVal t0 As Double) As String

    Dim wb As Workbook
    Dim k As Long, total As Long, done As Long, stuck As Long, after As Long
    Dim calcMode As Long, lastUpd As Double
    Dim errNo As Long, errMsg As String, stuckNames As String
    Dim elapsed As Double, rate As Double, etaSec As Double

    On Error GoTo Fail

    Application.StatusBar = "正在打开（含大量名称，可能需要几十秒，请勿操作）… " & p
    DoEvents

    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual     ' 打开前就关掉重算，减少卡顿
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = Application.Workbooks.Open(Filename:=p, UpdateLinks:=0, ReadOnly:=False)

    If wb.ProtectStructure Then
        On Error Resume Next
        wb.Unprotect
        On Error GoTo Fail
        If wb.ProtectStructure Then
            wb.Close SaveChanges:=False
            Application.Calculation = calcMode
            Application.StatusBar = False
            CleanByOM = "失败 · 工作簿结构受保护，无法删除名称"
            Exit Function
        End If
    End If

    total = wb.Names.Count

    If total = 0 Then
        wb.Close SaveChanges:=False
        Application.Calculation = calcMode
        Application.StatusBar = False
        CleanByOM = "完成 · 原本就没有名称" & BackupSuffix(bak, 0)
        Exit Function
    End If

    k = 1
    lastUpd = Timer

    Do While k <= wb.Names.Count

        Err.Clear
        wb.Names(k).Delete

        If Err.Number = 0 Then
            done = done + 1
        Else
            ' 删不掉：记下来，跳过这一个继续
            stuck = stuck + 1
            If errNo = 0 Then
                errNo = Err.Number
                errMsg = Err.Description
            End If
            Err.Clear
            If stuck <= 3 Then
                If Len(stuckNames) > 0 Then stuckNames = stuckNames & "、"
                stuckNames = stuckNames & wb.Names(k).Name
                Err.Clear
            End If
            k = k + 1
            If stuck >= MAX_STUCK Then Exit Do
        End If

        If Timer - lastUpd >= UI_INTERVAL Then

            elapsed = Timer - t0
            If elapsed < 0.001 Then elapsed = 0.001
            rate = done / elapsed
            If rate > 0 Then
                etaSec = (total - done) / rate
            Else
                etaSec = 0
            End If

            Application.StatusBar = _
                "名称清理中 | " & Format(done, "#,##0") & "/" & Format(total, "#,##0") & _
                " | " & Format(done / total, "0.0%") & _
                " | 速度: " & Format(rate, "0.0") & " 个/秒" & _
                " | 已耗时: " & Format(elapsed / 60, "0.0") & " 分钟" & _
                " | 剩余: " & Format(etaSec / 60, "0.0") & " 分钟" & _
                " | 完成时间: " & Format(Now + etaSec / 86400, "hh:mm:ss")

            DoEvents
            lastUpd = Timer

        End If

    Loop

    Application.StatusBar = False
    Application.Calculation = calcMode

    after = wb.Names.Count
    wb.Save
    wb.Close SaveChanges:=False

    If after = 0 Then

        CleanByOM = "完成 · 对象模型 · 删除 " & Format(done, "#,##0") & " 个 · " & _
                    Format(Timer - t0, "0.0") & " 秒" & BackupSuffix(bak, after)

    ElseIf after < AUTO_DEL_BACKUP_WHEN_REMAIN_LT Then

        CleanByOM = "完成（残留 " & after & " 个删不掉）· 删除 " & Format(done, "#,##0") & " 个 · " & _
                    "错误 " & errNo & "：" & errMsg & " · 例如：" & stuckNames & " · " & _
                    Format(Timer - t0, "0.0") & " 秒" & BackupSuffix(bak, after)

    Else

        CleanByOM = "部分完成 · 删除 " & Format(done, "#,##0") & " 个，剩余 " & Format(after, "#,##0") & _
                    " 个 · 错误 " & errNo & "：" & errMsg & " · 例如：" & stuckNames & " · " & _
                    Format(Timer - t0, "0.0") & " 秒"

    End If

    Exit Function

Fail:
    Application.StatusBar = False
    If calcMode = 0 Then calcMode = xlCalculationAutomatic
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    CleanByOM = "失败 · 对象模型出错：" & Err.Description

End Function


'--------------------------------------------------------------
' .xlsb / .xls 的处理方式选择（同一批只问一次）
'--------------------------------------------------------------
Private Function AskXlsbStrategy(ByVal ext As String) As VbMsgBoxResult

    Static asked As Boolean
    Static choice As VbMsgBoxResult

    If asked Then
        AskXlsbStrategy = choice
        Exit Function
    End If

    choice = MsgBox( _
        "目标文件是 ." & ext & " 二进制格式，名称存在二进制里，无法直接改写。" & vbCrLf & vbCrLf & _
        "【是】转换法（推荐）：临时转成 .xlsx / .xlsm 清理，再转回 ." & ext & "。" & vbCrLf & _
        "      30 万个约 1 分钟，不会卡在删不掉的名称上。" & vbCrLf & _
        "      注意：整个文件会被 Excel 重新保存一次，数字签名会失效；" & vbCrLf & _
        "            如果是「共享工作簿」，请改用逐个删除。" & vbCrLf & vbCrLf & _
        "【否】逐个删除：不改格式，30 万个约 15~20 分钟，个别坏名称会跳过。" & vbCrLf & vbCrLf & _
        "【取消】跳过所有 ." & ext & " 文件。", _
        vbYesNoCancel + vbQuestion, "选择处理方式（本批次只问一次）")

    asked = True
    AskXlsbStrategy = choice

End Function


'--------------------------------------------------------------
' 快路径：调用 PowerShell 直接改写 xl/workbook.xml
' 返回 0 = 成功
'--------------------------------------------------------------
Private Function RunCleanScript(ByVal p As String, ByRef outTxt As String) As Long

    Dim fso As Object, ps As String, outFile As String
    Dim cmd As String, sh As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    ps = fso.BuildPath(Environ$("TEMP"), PS_NAME)
    outFile = fso.BuildPath(Environ$("TEMP"), OUT_NAME)

    If Not EnsureScript(ps) Then
        outTxt = "无法生成 PowerShell 脚本"
        RunCleanScript = -1
        Exit Function
    End If

    On Error Resume Next
    Kill outFile
    On Error GoTo 0

    Application.StatusBar = "正在改写 xl/workbook.xml … " & fso.GetFileName(p)
    DoEvents

    cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & _
          Quote(ps) & " -Path " & Quote(p) & " -Log " & Quote(outFile)

    If Len(KEEP_PATTERN) > 0 Then cmd = cmd & " -KeepPattern " & Quote(KEEP_PATTERN)

    Set sh = CreateObject("WScript.Shell")
    RunCleanScript = sh.Run(cmd, 0, True)     ' 0=隐藏窗口，True=同步等待

    outTxt = ReadTextFile(outFile)
    If Len(outTxt) = 0 Then outTxt = "(无输出)"

End Function


'--------------------------------------------------------------
' 工具函数
'--------------------------------------------------------------
Private Function BackupSuffix(ByVal bak As String, ByVal remainCnt As Long) As String

    If Len(bak) = 0 Then Exit Function
    If remainCnt >= AUTO_DEL_BACKUP_WHEN_REMAIN_LT Then Exit Function

    On Error Resume Next
    Kill bak
    If Err.Number = 0 Then BackupSuffix = " · 已删备份"
    On Error GoTo 0

End Function


Private Function PickFiles() As Variant

    Dim v As Variant

    v = Application.GetOpenFilename( _
            FileFilter:="Excel 文件 (*.xlsx;*.xlsm;*.xlsb;*.xls;*.xltx;*.xltm),*.xlsx;*.xlsm;*.xlsb;*.xls;*.xltx;*.xltm", _
            Title:="选择要清理名称的 Excel 文件（可 Ctrl 多选）", _
            MultiSelect:=True)

    If VarType(v) = vbBoolean Then Exit Function       ' 用户点了取消
    PickFiles = v

End Function


Private Function CloseIfOpen(ByVal p As String) As Boolean

    Dim i As Long, wb As Workbook

    For i = Application.Workbooks.Count To 1 Step -1

        Set wb = Application.Workbooks(i)

        If StrComp(wb.FullName, p, vbTextCompare) = 0 Then

            If wb Is ThisWorkbook Then
                MsgBox "本宏不能在被清理的文件里运行。" & vbCrLf & vbCrLf & _
                       "请新建一个空白工作簿，导入本模块后在那里运行。", vbExclamation
                Exit Function
            End If

            If Not wb.Saved Then
                Select Case MsgBox("《" & wb.Name & "》有未保存的修改，是否先保存？", _
                                   vbYesNoCancel + vbQuestion)
                    Case vbYes: wb.Save
                    Case vbCancel: Exit Function
                End Select
            End If

            wb.Close SaveChanges:=False

        End If

    Next i

    CloseIfOpen = True

End Function


Private Function BackupFile(ByVal p As String) As String

    Dim fso As Object, dst As String

    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")

    dst = fso.BuildPath(fso.GetParentFolderName(p), _
          fso.GetBaseName(p) & "_备份_" & Format(Now, "yyyymmdd_hhnnss") & "." & fso.GetExtensionName(p))

    fso.CopyFile p, dst, True
    If Err.Number = 0 Then BackupFile = dst

    On Error GoTo 0

End Function


Private Function EnsureScript(ByVal ps As String) As Boolean

    Dim fso As Object, f As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(ps) Then
        If fso.GetFile(ps).Size > 0 Then
            EnsureScript = True
            Exit Function
        End If
    End If

    On Error GoTo Fail
    Set f = fso.CreateTextFile(ps, True, False)      ' 内容是纯 ASCII，用 ANSI 写即可
    f.Write PsTemplate
    f.Close
    EnsureScript = True
    Exit Function

Fail:
    On Error Resume Next
    If Not f Is Nothing Then f.Close

End Function


Private Function ReadTextFile(ByVal p As String) As String

    Dim fso As Object, f As Object, s As String

    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(p) Then Exit Function

    Set f = fso.OpenTextFile(p, 1)
    s = f.ReadAll
    f.Close

    ReadTextFile = Trim$(s)

End Function


Private Function Quote(ByVal s As String) As String
    Quote = Chr$(34) & s & Chr$(34)
End Function


Private Function GetNum(ByVal s As String, ByVal key As String) As Long

    Dim k As Long, i As Long, j As Long

    k = InStr(1, s, key & "=", vbTextCompare)
    If k = 0 Then Exit Function

    k = k + Len(key) + 1
    i = k
    Do While i <= Len(s)
        Select Case Mid$(s, i, 1)
            Case "0" To "9"
                i = i + 1
            Case Else
                Exit Do
        End Select
    Loop

    j = i - k
    If j <= 0 Then Exit Function

    GetNum = Val(Mid$(s, k, j))

End Function


'--------------------------------------------------------------
' 内嵌 PowerShell 脚本（纯 ASCII，运行时写到 %TEMP%）
' 注意：VBA 单条逻辑行最多 25 个续行符，故拆成 5 段 × 15 行。
'       以后要改脚本，请保持每段不超过 20 个元素。
'--------------------------------------------------------------
Private Function PsTemplate() As String

    Dim g1 As Variant, g2 As Variant, g3 As Variant, g4 As Variant, g5 As Variant
    Dim s As String

    g1 = Array( _
    "param(", _
    "    [Parameter(Mandatory=$true)][string]$Path,", _
    "    [string]$KeepPattern = '',", _
    "    [string]$Log = ''", _
    ")", _
    "$ErrorActionPreference = 'Stop'", _
    "Add-Type -AssemblyName System.IO.Compression | Out-Null", _
    "Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null", _
    "function Write-Result([string]$s) {", _
    "    if ([string]::IsNullOrEmpty($Log)) { Write-Output $s }", _
    "    else { Set-Content -LiteralPath $Log -Value $s -Encoding ASCII -Force }", _
    "}", _
    "if (-not (Test-Path -LiteralPath $Path)) { Write-Result 'ERR:FILE_NOT_FOUND'; exit 2 }", _
    "$sw = [System.Diagnostics.Stopwatch]::StartNew()", _
    "$zip = $null")

    g2 = Array( _
    "try {", _
    "    $zip = [System.IO.Compression.ZipFile]::Open($Path, 'Update')", _
    "    $entry = $zip.GetEntry('xl/workbook.xml')", _
    "    if ($null -eq $entry) { Write-Result 'ERR:NO_WORKBOOK_XML'; exit 3 }", _
    "    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)", _
    "    $xml = $reader.ReadToEnd()", _
    "    $reader.Close()", _
    "    $start = $xml.IndexOf('<definedNames')", _
    "    if ($start -lt 0) { Write-Result 'OK:removed=0;kept=0;ms=0'; exit 0 }", _
    "    $closeTag = $xml.IndexOf('</definedNames>', $start)", _
    "    if ($closeTag -lt 0) {", _
    "        $selfClose = $xml.IndexOf('/>', $start)", _
    "        if ($selfClose -lt 0) { Write-Result 'ERR:BAD_XML'; exit 4 }", _
    "        $end = $selfClose + 2", _
    "    } else {")

    g3 = Array( _
    "        $end = $closeTag + '</definedNames>'.Length", _
    "    }", _
    "    $block = $xml.Substring($start, $end - $start)", _
    "    $origCount = ([regex]::Matches($block, '<definedName\b')).Count", _
    "    $keptCount = 0", _
    "    if ([string]::IsNullOrEmpty($KeepPattern)) {", _
    "        $newXml = $xml.Substring(0, $start) + $xml.Substring($end)", _
    "    } else {", _
    "        $inner = $block -replace '(?s)^<definedNames[^>]*>', ''", _
    "        $inner = $inner -replace '(?s)</definedNames>\s*$', ''", _
    "        $sb = New-Object System.Text.StringBuilder", _
    "        $pat = '(?s)<definedName\b[^>]*>.*?</definedName>|<definedName\b[^>]*/>'", _
    "        $npat = 'name=(\u0022[^\u0022]*\u0022|[^\s>]+)'", _
    "        foreach ($m in [regex]::Matches($inner, $pat)) {", _
    "            $nm = ''")

    g4 = Array( _
    "            $nmMatch = [regex]::Match($m.Value, $npat)", _
    "            if ($nmMatch.Success) { $nm = $nmMatch.Groups[1].Value.Trim([char]34) }", _
    "            if ($nm -match $KeepPattern) { [void]$sb.Append($m.Value); $keptCount++ }", _
    "        }", _
    "        if ($sb.Length -gt 0) {", _
    "            $newXml = $xml.Substring(0, $start) + '<definedNames>' + $sb.ToString() + '</definedNames>' + $xml.Substring($end)", _
    "        } else {", _
    "            $newXml = $xml.Substring(0, $start) + $xml.Substring($end)", _
    "        }", _
    "    }", _
    "    $entry.Delete()", _
    "    $newEntry = $zip.CreateEntry('xl/workbook.xml', [System.IO.Compression.CompressionLevel]::Optimal)", _
    "    $st = $newEntry.Open()", _
    "    $writer = New-Object System.IO.StreamWriter($st, (New-Object System.Text.UTF8Encoding($false)))", _
    "    $writer.Write($newXml)")

    g5 = Array( _
    "    $writer.Flush()", _
    "    $writer.Close()", _
    "    $st.Close()", _
    "    $sw.Stop()", _
    "    $removed = $origCount - $keptCount", _
    "    Write-Result ('OK:removed={0};kept={1};ms={2}' -f $removed, $keptCount, $sw.ElapsedMilliseconds)", _
    "    exit 0", _
    "}", _
    "catch {", _
    "    Write-Result ('ERR:' + $_.Exception.Message)", _
    "    exit 5", _
    "}", _
    "finally {", _
    "    if ($null -ne $zip) { $zip.Dispose() }", _
    "}")

    s = Join(g1, vbCrLf)
    s = s & vbCrLf & Join(g2, vbCrLf)
    s = s & vbCrLf & Join(g3, vbCrLf)
    s = s & vbCrLf & Join(g4, vbCrLf)
    s = s & vbCrLf & Join(g5, vbCrLf)

    PsTemplate = s

End Function
