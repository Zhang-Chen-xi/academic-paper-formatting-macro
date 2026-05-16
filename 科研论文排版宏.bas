' ============================================================
' 科研论文排版 宏
' 作者：张辰熙 & Vibe Coding Agents
' 联系方式：chenxizhang_@outlook.com
' 第一版日期：2026年4月26日 | 最近修改：2026年5月16日
' 许可证：Apache License 2.0
' ============================================================
' 功能：对选中的中英文混合文本统一应用：
'   1、段落缩进（左缩进 + 段首额外缩进，见下方常量区）
'   2、中文 → 宋体 小四（12磅）
'   3、英文/数字/希腊字母/科研符号 → Times New Roman 小四（12磅）
'   4、行距：单倍行距，段前段后 0 磅
'   5、智能引号/括号字符自动转换（根据内容语言切换中英文标点）
' 推荐快捷键：Ctrl + Alt + J
' ============================================================

' ╔══════════════════════════════════════════════════════════╗
' ║  ★ 个性化常量区 ★ —— 修改这里即可适配不同排版规范    ║
' ╚══════════════════════════════════════════════════════════╝
Private Const INDENT_LEFT As Integer = 2          ' 整段左缩进（字符数）
Private Const INDENT_FIRSTLINE As Integer = 2     ' 段首额外缩进（字符数）
Private Const FONT_CJK As String = "宋体"         ' 中文字体
Private Const FONT_LATIN As String = "Times New Roman"  ' 西文字体
Private Const FONT_SIZE As Single = 12            ' 字号（磅），12磅 = 小四

Sub 科研论文排版()
    ' 核心优化：1:1复刻手动字体设置逻辑，完美区分中文/英文/希腊字母/科研符号
    ' 附加优化：关闭屏幕刷新提速 + 完整错误处理 + 格式保护
    ' 新增：智能引号/括号字符自动转换（根据内容语言自动切换中英文标点）
    On Error GoTo ErrorHandler
    
    ' 1. 校验是否选中文本
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        MsgBox "请先选中（拖选高亮）需要排版的段落文本，再运行此宏。" & vbCrLf & _
               "提示：将光标放在段落中但未拖选文字时，宏无法确定作用范围。", _
               vbExclamation, "科研论文排版"
        Exit Sub
    End If
    
    ' 2. 关闭屏幕刷新（大幅提升长文本排版速度，无卡顿）
    Application.ScreenUpdating = False
    
    ' --------------------------
    ' 【第一步】段落缩进设置
    ' --------------------------
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = INDENT_LEFT
        .CharacterUnitFirstLineIndent = INDENT_FIRSTLINE
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
    
    ' --------------------------
    ' 【第二步】字体设置（完全复现手动操作逻辑）
    ' --------------------------
    With Selection.Font
        ' 第一步：全局设中文字体 → Word仅对CJK字符生效
        .Name = FONT_CJK
        ' 第二步：全局设西文字体 → Word仅对非CJK字符生效，中文保持不变
        .Name = FONT_LATIN
        ' 统一字号
        .Size = FONT_SIZE
    End With
    
    ' --------------------------
    ' 【第三步】智能标点字符转换 + 字体校正
    ' --------------------------
    Call FixQuoteFonts(Selection.Range)
    Call FixBracketFonts(Selection.Range)
    Call FixDashAndEllipsisFonts(Selection.Range)
    
ExitHandler:
    ' 无论是否报错，都强制恢复屏幕刷新，避免Word界面卡死
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "排版出错：" & Err.Description, vbCritical, "错误提示"
    Resume ExitHandler
End Sub

' ==============================================
' 智能引号字体校正 + 字符转换
' ==============================================
Private Sub FixQuoteFonts(rng As Range)
    Call ProcessDoubleQuotes(rng)
    Call ProcessSingleQuotes(rng)
    Call ProcessCJKCornerQuotes(rng)
End Sub

' ==============================================
' 辅助：用 Find 在指定范围找 text1 或 text2 中最先出现的位置
' 未找到返回 -1
' ==============================================
Private Function FindNearestPos(doc As Document, startPos As Long, endPos As Long, text1 As String, text2 As String) As Long
    Dim pos1 As Long: pos1 = -1
    Dim pos2 As Long: pos2 = -1
    Dim fr As Range
    
    If startPos >= endPos Then
        FindNearestPos = -1
        Exit Function
    End If
    
    Set fr = doc.Range(startPos, endPos)
    With fr.Find
        .ClearFormatting
        .text = text1
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    If fr.Find.Execute Then pos1 = fr.Start
    
    Set fr = doc.Range(startPos, endPos)
    With fr.Find
        .ClearFormatting
        .text = text2
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    If fr.Find.Execute Then pos2 = fr.Start
    
    If pos1 = -1 And pos2 = -1 Then
        FindNearestPos = -1
    ElseIf pos1 = -1 Then
        FindNearestPos = pos2
    ElseIf pos2 = -1 Then
        FindNearestPos = pos1
    ElseIf pos1 <= pos2 Then
        FindNearestPos = pos1
    Else
        FindNearestPos = pos2
    End If
End Function

' ==============================================
' 双引号：" (U+0022) 与 ""(U+201C/U+201D) 互转
' 中文内容→弯引号+宋体，英文内容→直引号+TNR
' ==============================================
Private Sub ProcessDoubleQuotes(rng As Range)
    Dim pos As Long
    Dim openPos As Long, closePos As Long
    Dim contentRange As Range
    Dim chineseRatio As Double
    Dim openRange As Range, closeRange As Range
    
    Dim CHN_OPEN As String: CHN_OPEN = ChrW(&H201C)
    Dim CHN_CLOSE As String: CHN_CLOSE = ChrW(&H201D)
    Dim ENG_QUOTE As String: ENG_QUOTE = ChrW(&H22)
    
    pos = rng.Start
    
    Do While pos < rng.End
        openPos = FindNearestPos(rng.Document, pos, rng.End, CHN_OPEN, ENG_QUOTE)
        If openPos = -1 Then Exit Do
        
        closePos = FindNearestPos(rng.Document, openPos + 1, rng.End, CHN_CLOSE, ENG_QUOTE)
        If closePos = -1 Then Exit Do
        
        If closePos > openPos + 1 Then
            Set contentRange = rng.Document.Range(openPos + 1, closePos)
            
            If contentRange.InlineShapes.Count = 0 And contentRange.Fields.Count = 0 Then
                chineseRatio = GetChineseRatio(contentRange)
                
                If chineseRatio >= 0.5 Then
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.text = CHN_CLOSE
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.Font.Name = FONT_CJK
                    closeRange.Font.Size = FONT_SIZE
                    
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.text = CHN_OPEN
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.Font.Name = FONT_CJK
                    openRange.Font.Size = FONT_SIZE
                Else
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.text = ENG_QUOTE
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.Font.Name = FONT_LATIN
                    closeRange.Font.Size = FONT_SIZE
                    
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.text = ENG_QUOTE
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.Font.Name = FONT_LATIN
                    openRange.Font.Size = FONT_SIZE
                End If
            End If
        End If
        
        pos = closePos + 1
    Loop
End Sub

' ==============================================
' 单引号：' (U+0027) 与 ''(U+2018/U+2019) 互转
' ==============================================
Private Sub ProcessSingleQuotes(rng As Range)
    Dim pos As Long
    Dim openPos As Long, closePos As Long
    Dim contentRange As Range
    Dim chineseRatio As Double
    Dim openRange As Range, closeRange As Range
    
    Dim CHN_OPEN As String: CHN_OPEN = ChrW(&H2018)
    Dim CHN_CLOSE As String: CHN_CLOSE = ChrW(&H2019)
    Dim ENG_QUOTE As String: ENG_QUOTE = ChrW(&H27)
    
    pos = rng.Start
    
    Do While pos < rng.End
        openPos = FindNearestPos(rng.Document, pos, rng.End, CHN_OPEN, ENG_QUOTE)
        If openPos = -1 Then Exit Do
        
        closePos = FindNearestPos(rng.Document, openPos + 1, rng.End, CHN_CLOSE, ENG_QUOTE)
        If closePos = -1 Then Exit Do
        
        If closePos > openPos + 1 Then
            Set contentRange = rng.Document.Range(openPos + 1, closePos)
            
            If contentRange.InlineShapes.Count = 0 And contentRange.Fields.Count = 0 Then
                chineseRatio = GetChineseRatio(contentRange)
                
                If chineseRatio >= 0.5 Then
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.text = CHN_CLOSE
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.Font.Name = FONT_CJK
                    closeRange.Font.Size = FONT_SIZE
                    
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.text = CHN_OPEN
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.Font.Name = FONT_CJK
                    openRange.Font.Size = FONT_SIZE
                Else
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.text = ENG_QUOTE
                    Set closeRange = rng.Document.Range(closePos, closePos + 1)
                    closeRange.Font.Name = FONT_LATIN
                    closeRange.Font.Size = FONT_SIZE
                    
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.text = ENG_QUOTE
                    Set openRange = rng.Document.Range(openPos, openPos + 1)
                    openRange.Font.Name = FONT_LATIN
                    openRange.Font.Size = FONT_SIZE
                End If
            End If
        End If
        
        pos = closePos + 1
    Loop
End Sub

' ==============================================
' CJK角引号「」『』：仅字体校正，不转换字符
' ==============================================
Private Sub ProcessCJKCornerQuotes(rng As Range)
    Dim quotePairs As Variant
    Dim i As Integer
    Dim startPos As Long
    Dim findRange As Range
    Dim contentRange As Range
    Dim chineseRatio As Double
    
    quotePairs = Array(ChrW(&H300C), ChrW(&H300D), ChrW(&H300E), ChrW(&H300F))
    
    For i = LBound(quotePairs) To UBound(quotePairs) Step 2
        startPos = rng.Start
        
        Do
            Set findRange = rng.Document.Range(startPos, rng.End)
            If Not findRange.Find.Execute(FindText:=quotePairs(i), Forward:=True, Wrap:=wdFindStop) Then
                Exit Do
            End If
            
            Dim quoteStart As Long
            quoteStart = findRange.Start
            
            Set findRange = rng.Document.Range(quoteStart + 1, rng.End)
            If Not findRange.Find.Execute(FindText:=quotePairs(i + 1), Forward:=True, Wrap:=wdFindStop) Then
                Exit Do
            End If
            
            Dim quoteEnd As Long
            quoteEnd = findRange.End
            
            If quoteEnd > quoteStart + 2 Then
                Set contentRange = rng.Document.Range(quoteStart + 1, quoteEnd - 1)
                
                If contentRange.InlineShapes.Count = 0 And contentRange.Fields.Count = 0 Then
                    chineseRatio = GetChineseRatio(contentRange)
                    If chineseRatio >= 0.5 Then
                        rng.Document.Range(quoteStart, quoteStart + 1).Font.Name = FONT_CJK
                        rng.Document.Range(quoteEnd - 1, quoteEnd).Font.Name = FONT_CJK
                    Else
                        rng.Document.Range(quoteStart, quoteStart + 1).Font.Name = FONT_LATIN
                        rng.Document.Range(quoteEnd - 1, quoteEnd).Font.Name = FONT_LATIN
                    End If
                End If
            End If
            
            startPos = quoteEnd
        Loop
    Next i
End Sub

' ==============================================
' 智能括号字体校正 + 字符转换
' 等价组：( ) <-> （ ）、[ ] <-> 【 】、{ } <-> ｛ ｝
' 另处理：《》仅字体校正
' ==============================================
Private Sub FixBracketFonts(rng As Range)
    Call ProcessBracketGroup(rng, "(", ")", ChrW(&HFF08), ChrW(&HFF09))
    Call ProcessBracketGroup(rng, "[", "]", ChrW(&H3010), ChrW(&H3011))
    Call ProcessBracketGroup(rng, "{", "}", ChrW(&HFF5B), ChrW(&HFF5D))
    Call ProcessBookTitleBrackets(rng)
End Sub

' ==============================================
' 单组括号互转（Find定位开括号 + 逐字符匹配嵌套闭括号）
' ==============================================
Private Sub ProcessBracketGroup(rng As Range, engOpen As String, engClose As String, chnOpen As String, chnClose As String)
    Dim pos As Long
    Dim bracketStart As Long
    Dim currentPos As Long
    Dim depth As Integer
    Dim contentRange As Range
    Dim chineseRatio As Double
    Dim ch As String
    Dim bracketEnd As Long
    Dim openRange As Range, closeRange As Range
    
    pos = rng.Start
    
    Do While pos < rng.End
        bracketStart = FindNearestPos(rng.Document, pos, rng.End, engOpen, chnOpen)
        If bracketStart = -1 Then Exit Do
        
        ' 逐字符扫描找匹配闭括号（处理嵌套）
        currentPos = bracketStart + 1
        depth = 1
        
        Do While currentPos < rng.End And depth > 0
            On Error Resume Next
            ch = rng.Document.Range(currentPos, currentPos + 1).text
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                currentPos = currentPos + 1
                GoTo SkipChar
            End If
            On Error GoTo 0
            
            If Len(ch) > 0 Then
                If ch = engOpen Or ch = chnOpen Then
                    depth = depth + 1
                ElseIf ch = engClose Or ch = chnClose Then
                    depth = depth - 1
                End If
            End If
            
            If depth > 0 Then currentPos = currentPos + 1
SkipChar:
        Loop
        
        If depth = 0 Then
            bracketEnd = currentPos + 1
            
            If bracketEnd > bracketStart + 2 Then
                Set contentRange = rng.Document.Range(bracketStart + 1, bracketEnd - 1)
                
                If contentRange.InlineShapes.Count = 0 And contentRange.Fields.Count = 0 Then
                    chineseRatio = GetChineseRatio(contentRange)
                    
                    If chineseRatio >= 0.5 Then
                        Set closeRange = rng.Document.Range(currentPos, currentPos + 1)
                        closeRange.text = chnClose
                        Set closeRange = rng.Document.Range(currentPos, currentPos + 1)
                        closeRange.Font.Name = FONT_CJK
                        closeRange.Font.Size = FONT_SIZE
                        
                        Set openRange = rng.Document.Range(bracketStart, bracketStart + 1)
                        openRange.text = chnOpen
                        Set openRange = rng.Document.Range(bracketStart, bracketStart + 1)
                        openRange.Font.Name = FONT_CJK
                        openRange.Font.Size = FONT_SIZE
                    Else
                        Set closeRange = rng.Document.Range(currentPos, currentPos + 1)
                        closeRange.text = engClose
                        Set closeRange = rng.Document.Range(currentPos, currentPos + 1)
                        closeRange.Font.Name = FONT_LATIN
                        closeRange.Font.Size = FONT_SIZE
                        
                        Set openRange = rng.Document.Range(bracketStart, bracketStart + 1)
                        openRange.text = engOpen
                        Set openRange = rng.Document.Range(bracketStart, bracketStart + 1)
                        openRange.Font.Name = FONT_LATIN
                        openRange.Font.Size = FONT_SIZE
                    End If
                End If
            End If
            
            pos = bracketEnd
        Else
            pos = bracketStart + 1
        End If
    Loop
End Sub

' ==============================================
' 《》书名号：仅字体校正
' ==============================================
Private Sub ProcessBookTitleBrackets(rng As Range)
    Dim startPos As Long
    Dim findRange As Range
    Dim contentRange As Range
    Dim chineseRatio As Double
    Dim bracketOpen As String: bracketOpen = ChrW(&H300A)
    Dim bracketClose As String: bracketClose = ChrW(&H300B)
    Dim bracketStart As Long, bracketEnd As Long
    
    startPos = rng.Start
    
    Do
        Set findRange = rng.Document.Range(startPos, rng.End)
        If Not findRange.Find.Execute(FindText:=bracketOpen, Forward:=True, Wrap:=wdFindStop) Then Exit Do
        bracketStart = findRange.Start
        
        Set findRange = rng.Document.Range(bracketStart + 1, rng.End)
        If Not findRange.Find.Execute(FindText:=bracketClose, Forward:=True, Wrap:=wdFindStop) Then Exit Do
        bracketEnd = findRange.End
        
        If bracketEnd > bracketStart + 2 Then
            Set contentRange = rng.Document.Range(bracketStart + 1, bracketEnd - 1)
            If contentRange.InlineShapes.Count = 0 And contentRange.Fields.Count = 0 Then
                chineseRatio = GetChineseRatio(contentRange)
                If chineseRatio >= 0.5 Then
                    rng.Document.Range(bracketStart, bracketStart + 1).Font.Name = FONT_CJK
                    rng.Document.Range(bracketEnd - 1, bracketEnd).Font.Name = FONT_CJK
                Else
                    rng.Document.Range(bracketStart, bracketStart + 1).Font.Name = FONT_LATIN
                    rng.Document.Range(bracketEnd - 1, bracketEnd).Font.Name = FONT_LATIN
                End If
            End If
        End If
        
        startPos = bracketEnd
    Loop
End Sub

' ==============================================
' 破折号和省略号字体校正
' ==============================================
Private Sub FixDashAndEllipsisFonts(rng As Range)
    Dim chinesePunctuations As Variant
    Dim englishPunctuations As Variant
    Dim i As Integer
    Dim findRange As Range
    
    chinesePunctuations = Array(ChrW(&H2014), ChrW(&H2026))
    englishPunctuations = Array(ChrW(&H2013), ChrW(&H2014), "...")
    
    For i = LBound(chinesePunctuations) To UBound(chinesePunctuations)
        Set findRange = rng.Duplicate
        With findRange.Find
            .ClearFormatting
            .text = chinesePunctuations(i)
            .Replacement.ClearFormatting
            .Replacement.Font.Name = FONT_CJK
            .Execute Replace:=wdReplaceAll, Wrap:=wdFindStop
        End With
    Next i
    
    For i = LBound(englishPunctuations) To UBound(englishPunctuations)
        Set findRange = rng.Duplicate
        With findRange.Find
            .ClearFormatting
            .text = englishPunctuations(i)
            .Replacement.ClearFormatting
            .Replacement.Font.Name = FONT_LATIN
            .Execute Replace:=wdReplaceAll, Wrap:=wdFindStop
        End With
    Next i
End Sub

' ==============================================
' 辅助：计算文本范围内的中文占比
' 仅统计汉字和英文字母，排除空格/数字/标点干扰
' ==============================================
Private Function GetChineseRatio(rng As Range) As Double
    Dim totalChars As Long
    Dim chineseChars As Long
    Dim englishChars As Long
    Dim i As Long
    Dim charCode As Long
    Dim charText As String
    
    totalChars = rng.Characters.Count
    chineseChars = 0
    englishChars = 0
    
    If totalChars = 0 Then
        GetChineseRatio = 1#
        Exit Function
    End If
    
    For i = 1 To totalChars
        charText = rng.Characters(i).text
        If Len(charText) = 0 Then GoTo NextChar
        
        charCode = AscW(charText)
        
        If charCode >= &H4E00 And charCode <= &H9FFF Then
            chineseChars = chineseChars + 1
        ElseIf (charCode >= 65 And charCode <= 90) Or (charCode >= 97 And charCode <= 122) Then
            englishChars = englishChars + 1
        End If
NextChar:
    Next i
    
    If chineseChars + englishChars = 0 Then
        GetChineseRatio = 1#
        Exit Function
    End If
    
    GetChineseRatio = chineseChars / (chineseChars + englishChars)
End Function


' ╔══════════════════════════════════════════════════════════════════╗
' ║  以下为旧版存档代码，仅供参考回溯，不影响上方主程序运行。      ║
' ║  如需使用旧版逻辑（无智能标点功能），可运行此 Sub。            ║
' ╚══════════════════════════════════════════════════════════════════╝
Sub 科研论文排版_存档()
    '第一版日期：2026年4月26日
    '功能：基础排版（缩进 + 字体 + 字号），不含智能标点转换
    On Error GoTo ErrorHandler
    
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        MsgBox "请先选中（拖选高亮）需要排版的段落文本，再运行此宏。" & vbCrLf & _
               "提示：将光标放在段落中但未拖选文字时，宏无法确定作用范围。", _
               vbExclamation, "科研论文排版"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = INDENT_LEFT
        .CharacterUnitFirstLineIndent = INDENT_FIRSTLINE
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
    
    With Selection.Font
        .Name = FONT_CJK
        .Name = FONT_LATIN
        .Size = FONT_SIZE
    End With

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "排版出错：" & Err.Description, vbCritical, "错误提示"
    Resume ExitHandler
End Sub
