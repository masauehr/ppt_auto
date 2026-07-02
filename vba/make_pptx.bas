Attribute VB_Name = "MakePPTX"
'=============================================================
' PowerPoint定型スライド自動生成マクロ
'
' 収録マクロ:
'   MakePPTX             - コード内に直書きした雛形からスライドを生成
'   MakePPTXFromMarkdown - Markdownファイルを読み込んでスライドを生成
'                           （表・コードブロック・タイトル自動取得に対応。
'                           `>`引用・`![]()`画像は非対応）
'
' 使い方: PowerPointのVBAエディタ（Windows: Alt+F11 / Mac: Option+F11）
'         からいずれかのマクロを実行
'=============================================================

Option Explicit

' ---- 設定エリア（ここを編集してスライド内容を変更する） ----

Private Const PRES_TITLE    As String = "プレゼンテーションタイトル"
Private Const PRES_SUBTITLE As String = "サブタイトル"
Private Const PRES_AUTHOR   As String = "作成者名"
Private Const PRES_DATE     As String = "2026年4月29日"

' スライド数（SLIDE_COUNT に合わせてSlideTitles・SlideBodies配列を編集する）
Private Const SLIDE_COUNT As Integer = 4

' ---- カラー定義 ----
Private Const COLOR_ACCENT  As Long = &H7D491F  ' BGRで格納（VBA仕様）→ 実際は #1F497D
Private Const COLOR_WHITE   As Long = &HFFFFFF
Private Const COLOR_DARK    As Long = &H262626
Private Const COLOR_LIGHT   As Long = &HF2F2F2
Private Const COLOR_CODE_BG As Long = &HE8E8E8  ' コードブロック背景（本文背景より少し濃いグレー、RGB均等なのでBGR変換の影響なし）

' ----------------------------------------------------------------
' ユーティリティ: センチメートル→ポイント変換
' （Mac版PowerPointのVBAには Application.CentimetersToPoints が実装され
' ていないため、計算式で代替する。1cm = 72/2.54 pt）
' ----------------------------------------------------------------
Private Function CmToPt(cm As Single) As Single
    CmToPt = cm * 28.346456692913385
End Function

'=============================================================
' メインプロシージャ
'=============================================================
Public Sub MakePPTX()

    Dim prs As Presentation
    Dim titles(1 To SLIDE_COUNT) As String
    Dim bodies(1 To SLIDE_COUNT) As String
    Dim types(1 To SLIDE_COUNT)  As String

    ' ---- スライド内容を設定する ----
    titles(1) = "はじめに"
    bodies(1) = "ここに本文テキストを入力します。" & Chr(13) & _
                "箇条書きも改行で区切って記述できます。"
    types(1)  = "content"

    titles(2) = "内容1"
    bodies(2) = "1つ目のセクションの内容です。"
    types(2)  = "content"

    titles(3) = "内容2"
    bodies(3) = "2つ目のセクションの内容です。"
    types(3)  = "content"

    titles(4) = "まとめ"
    bodies(4) = "まとめのテキストをここに記述します。"
    types(4)  = "summary"
    ' --------------------------------

    ' 新規プレゼンテーション作成
    Set prs = Presentations.Add(WithWindow:=True)

    ' スライドサイズを16:9に設定
    With prs.PageSetup
        .SlideWidth  = CmToPt(33.867)
        .SlideHeight = CmToPt(19.05)
    End With

    ' 表紙スライド
    Call MakeTitleSlide(prs, PRES_TITLE, PRES_SUBTITLE, PRES_AUTHOR, PRES_DATE)

    ' 目次スライド
    If SLIDE_COUNT >= 2 Then
        Call MakeTocSlide(prs, titles, SLIDE_COUNT)
    End If

    ' 本文・まとめスライド
    Dim i As Integer
    For i = 1 To SLIDE_COUNT
        If types(i) = "summary" Then
            Call MakeSummarySlide(prs, titles(i), bodies(i))
        Else
            Call MakeContentSlide(prs, titles(i), bodies(i), New Collection, New Collection)
        End If
    Next i

    MsgBox "スライド生成完了！（" & prs.Slides.Count & "枚）", vbInformation

End Sub

'=============================================================
' Markdownファイルからスライドを生成するメインプロシージャ
'=============================================================
Public Sub MakePPTXFromMarkdown()

    Dim mdPath As String
    mdPath = PickMarkdownFile()
    If mdPath = "" Then Exit Sub

    Dim mdText As String
    mdText = ReadTextFile(mdPath)

    Dim cfgTitle As String, cfgSubtitle As String
    Dim cfgAuthor As String, cfgDate As String
    Dim bodyText As String
    bodyText = ParseFrontMatter(mdText, cfgTitle, cfgSubtitle, cfgAuthor, cfgDate)

    ' titleが未指定の場合、本文の見出しから自動取得する
    ' （最初の"# "見出し、なければ最初の"## "見出しをタイトルとして使う）
    If Trim(cfgTitle) = "" Then
        cfgTitle = ExtractAutoTitle(bodyText)
    End If

    Dim slideTitles() As String
    Dim slideBodies() As String
    Dim slideTypes() As String
    Dim slideTables() As Collection
    Dim slideCodeBlocks() As Collection
    Dim slideCount As Integer
    Call ParseSlides(bodyText, slideTitles, slideBodies, slideTypes, slideTables, slideCodeBlocks, slideCount)

    If slideCount = 0 Then
        MsgBox "スライドが見つかりませんでした（## 見出しのスライドが1枚もありません）", vbExclamation
        Exit Sub
    End If

    Dim prs As Presentation
    Set prs = Presentations.Add(WithWindow:=True)

    With prs.PageSetup
        .SlideWidth  = CmToPt(33.867)
        .SlideHeight = CmToPt(19.05)
    End With

    Call MakeTitleSlide(prs, cfgTitle, cfgSubtitle, cfgAuthor, cfgDate)

    If slideCount >= 2 Then
        Call MakeTocSlide(prs, slideTitles, slideCount)
    End If

    Dim i As Integer
    For i = 1 To slideCount
        If slideTypes(i) = "summary" Then
            Call MakeSummarySlide(prs, slideTitles(i), slideBodies(i))
        Else
            Call MakeContentSlide(prs, slideTitles(i), slideBodies(i), slideTables(i), slideCodeBlocks(i))
        End If
    Next i

    MsgBox "スライド生成完了！（" & prs.Slides.Count & "枚）" & Chr(13) & _
           "元ファイル: " & mdPath, vbInformation

End Sub

'=============================================================
' ユーティリティ: Markdownファイルのパスを取得する
' （Mac版VBAではApplication.FileDialogが使えないためInputBoxで代用）
'=============================================================
Private Function PickMarkdownFile() As String

    Dim p As String
    p = InputBox("読み込むMarkdownファイルのフルパスを入力してください" & Chr(13) & _
                 "（例: /Users/username/projects/ppt_auto/sample_slides.md）", _
                 "Markdownファイルを選択")

    If Trim(p) = "" Then
        PickMarkdownFile = ""
        Exit Function
    End If

    p = Trim(p)
    If Dir(p) = "" Then
        MsgBox "ファイルが見つかりません: " & p, vbExclamation
        PickMarkdownFile = ""
        Exit Function
    End If

    PickMarkdownFile = p

End Function

'=============================================================
' ユーティリティ: テキストファイルをUTF-8として読み込む
' （Mac版VBAの Open...For Input / Line Input はUTF-8のマルチバイト文字を
' 正しく扱えず文字化けするため、バイナリで読み込んで自前でUTF-8デコードする）
'=============================================================
Private Function ReadTextFile(ByVal path As String) As String

    Dim fnum As Integer
    fnum = FreeFile

    Dim fileLen As Long
    Open path For Binary Access Read As #fnum
    fileLen = LOF(fnum)

    If fileLen = 0 Then
        Close #fnum
        ReadTextFile = ""
        Exit Function
    End If

    Dim bytes() As Byte
    ReDim bytes(1 To fileLen)
    Get #fnum, 1, bytes
    Close #fnum

    ' UTF-8のBOM（EF BB BF）が付いていればスキップする
    Dim startPos As Long
    startPos = 1
    If fileLen >= 3 Then
        If bytes(1) = &HEF And bytes(2) = &HBB And bytes(3) = &HBF Then
            startPos = 4
        End If
    End If

    Dim raw As String
    raw = Utf8BytesToString(bytes, startPos)

    ' 改行コードを vbLf に正規化する（CRLF/CRいずれにも対応）
    raw = Replace(raw, Chr(13) & Chr(10), Chr(10))
    raw = Replace(raw, Chr(13), Chr(10))

    ReadTextFile = raw

End Function

'=============================================================
' ユーティリティ: バイト配列をUTF-8としてデコードし文字列に変換する
'=============================================================
Private Function Utf8BytesToString(bytes() As Byte, ByVal startAt As Long) As String

    Dim result As String
    Dim hi As Long
    hi = UBound(bytes)

    Dim i As Long
    i = startAt

    Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long
    Dim codepoint As Long

    Do While i <= hi
        b1 = bytes(i)
        If b1 < &H80 Then
            codepoint = b1
            i = i + 1
        ElseIf (b1 And &HE0) = &HC0 And i + 1 <= hi Then
            b2 = bytes(i + 1)
            codepoint = ((b1 And &H1F) * &H40) + (b2 And &H3F)
            i = i + 2
        ElseIf (b1 And &HF0) = &HE0 And i + 2 <= hi Then
            b2 = bytes(i + 1)
            b3 = bytes(i + 2)
            codepoint = ((b1 And &HF) * &H1000) + ((b2 And &H3F) * &H40) + (b3 And &H3F)
            i = i + 3
        ElseIf (b1 And &HF8) = &HF0 And i + 3 <= hi Then
            b2 = bytes(i + 1)
            b3 = bytes(i + 2)
            b4 = bytes(i + 3)
            codepoint = ((b1 And &H7) * &H40000) + ((b2 And &H3F) * &H1000) + ((b3 And &H3F) * &H40) + (b4 And &H3F)
            i = i + 4
        Else
            codepoint = 63 ' 不正なバイト列は「?」に置き換える
            i = i + 1
        End If

        If codepoint > 65535 Then
            ' 16進リテラル &HFFFF 等はVBAでは符号付きIntegerとして解釈され
            ' 負数（-1）になってしまうため、10進数リテラルを使う
            Dim v As Long
            v = codepoint - 65536
            result = result & ChrW(55296 + (v \ 1024)) & ChrW(56320 + (v Mod 1024))
        Else
            result = result & ChrW(codepoint)
        End If
    Loop

    Utf8BytesToString = result

End Function

'=============================================================
' ユーティリティ: フロントマター（---～---）を解析する
' タイトル等をByRef引数に格納し、残りの本文を戻り値で返す
'=============================================================
Private Function ParseFrontMatter(ByVal src As String, _
                                   ByRef outTitle As String, ByRef outSubtitle As String, _
                                   ByRef outAuthor As String, ByRef outDate As String) As String

    outTitle = "": outSubtitle = "": outAuthor = "": outDate = ""

    Dim lines() As String
    lines = Split(src, vbLf)

    If UBound(lines) < 0 Then
        ParseFrontMatter = src
        Exit Function
    End If

    If Trim(lines(0)) <> "---" Then
        ParseFrontMatter = src
        Exit Function
    End If

    Dim i As Integer
    Dim endIdx As Integer
    endIdx = -1

    For i = 1 To UBound(lines)
        If Trim(lines(i)) = "---" Then
            endIdx = i
            Exit For
        End If

        Dim colonPos As Integer
        colonPos = InStr(lines(i), ":")
        If colonPos > 0 Then
            Dim k As String, v As String
            k = Trim(Left(lines(i), colonPos - 1))
            v = Trim(Mid(lines(i), colonPos + 1))
            Select Case k
                Case "title":    outTitle = v
                Case "subtitle": outSubtitle = v
                Case "author":   outAuthor = v
                Case "date":     outDate = v
            End Select
        End If
    Next i

    If endIdx = -1 Then
        ' 閉じの --- が見つからない場合はフロントマターなしとして扱う
        outTitle = "": outSubtitle = "": outAuthor = "": outDate = ""
        ParseFrontMatter = src
        Exit Function
    End If

    Dim rest As String
    rest = ""
    For i = endIdx + 1 To UBound(lines)
        rest = rest & lines(i) & vbLf
    Next i

    ParseFrontMatter = rest

End Function

'=============================================================
' ユーティリティ: フロントマターでtitleが省略された場合に、
' 本文の見出しから自動取得する（最初の"# "見出し、なければ
' 最初の"## "見出しを使う）
'=============================================================
Private Function ExtractAutoTitle(ByVal src As String) As String

    Dim lines() As String
    lines = Split(src, vbLf)

    Dim h2Title As String
    h2Title = ""

    Dim i As Long
    Dim ln As String
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Left(ln, 2) = "# " And Left(ln, 3) <> "## " Then
            ExtractAutoTitle = Trim(Mid(ln, 3))
            Exit Function
        ElseIf Left(ln, 3) = "## " And h2Title = "" Then
            h2Title = Trim(Mid(ln, 4))
        End If
    Next i

    ExtractAutoTitle = h2Title

End Function

'=============================================================
' ユーティリティ: "## " 見出しでスライドに分割する
' （各スライドの本文はParseSlideBodyでさらに本文/表/コードブロックに
' 分解し、outTables・outCodeBlocksにスライドごとのCollectionを格納する）
'=============================================================
Private Sub ParseSlides(ByVal src As String, _
                         ByRef outTitles() As String, ByRef outBodies() As String, _
                         ByRef outTypes() As String, ByRef outTables() As Collection, _
                         ByRef outCodeBlocks() As Collection, ByRef outCount As Integer)

    Dim lines() As String
    lines = Split(src, vbLf)

    Dim maxSlides As Integer
    maxSlides = UBound(lines) + 1
    If maxSlides < 1 Then maxSlides = 1

    ReDim outTitles(1 To maxSlides)
    ReDim outBodies(1 To maxSlides)
    ReDim outTypes(1 To maxSlides)
    ReDim outTables(1 To maxSlides)
    ReDim outCodeBlocks(1 To maxSlides)
    outCount = 0

    Dim curTitle As String
    Dim curBodyLines As String
    Dim curType As String
    Dim inSlide As Boolean
    inSlide = False

    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = lines(i)

        If Left(ln, 3) = "## " Then
            If inSlide Then
                outCount = outCount + 1
                outTitles(outCount) = curTitle
                Dim tbls As Collection, codes As Collection
                outBodies(outCount) = ParseSlideBody(curBodyLines, tbls, codes)
                Set outTables(outCount) = tbls
                Set outCodeBlocks(outCount) = codes
                outTypes(outCount) = curType
            End If

            Dim rawTitle As String
            rawTitle = Mid(ln, 4)
            curType = "content"
            If InStr(rawTitle, "<!-- summary -->") > 0 Then
                curType = "summary"
                rawTitle = Replace(rawTitle, "<!-- summary -->", "")
            End If
            curTitle = Trim(rawTitle)
            curBodyLines = ""
            inSlide = True
        ElseIf inSlide Then
            curBodyLines = curBodyLines & ln & vbLf
        End If
    Next i

    If inSlide Then
        outCount = outCount + 1
        outTitles(outCount) = curTitle
        Dim tbls2 As Collection, codes2 As Collection
        outBodies(outCount) = ParseSlideBody(curBodyLines, tbls2, codes2)
        Set outTables(outCount) = tbls2
        Set outCodeBlocks(outCount) = codes2
        outTypes(outCount) = curType
    End If

End Sub

'=============================================================
' ユーティリティ: スライド本文を本文テキスト・表・コードブロックに分解する
' （"- "箇条書きは"・"に変換してChr(13)区切りの本文にする。
' "|セル|セル|"の行は表としてoutTablesへ、"```"で囲まれた行は
' コードブロックとしてoutCodeBlocksへ、それぞれ集める）
'=============================================================
Private Function ParseSlideBody(ByVal raw As String, ByRef outTables As Collection, ByRef outCodeBlocks As Collection) As String

    Dim lines() As String
    Dim bodyResult As String
    Dim bodyFirst As Boolean
    Dim inCodeBlock As Boolean
    Dim codeLines As String
    Dim codeFirst As Boolean
    Dim tableHeader As Collection
    Dim tableRows As Collection
    Dim tablePendingHeader As Boolean
    Dim i As Long
    Dim ln As String
    Dim trimmed As String
    Dim cells As Collection
    Dim joined As String
    Dim cIdx As Long
    Dim outLine As String

    Set outTables = New Collection
    Set outCodeBlocks = New Collection

    lines = Split(raw, vbLf)

    bodyFirst = True
    inCodeBlock = False

    Set tableHeader = Nothing
    Set tableRows = Nothing
    tablePendingHeader = False

    For i = 0 To UBound(lines)
        ln = lines(i)
        trimmed = Trim(ln)

        ' コードフェンス（```）の開始・終了
        If Left(trimmed, 3) = "```" Then
            If inCodeBlock Then
                outCodeBlocks.Add codeLines
                inCodeBlock = False
            Else
                inCodeBlock = True
                codeLines = ""
                codeFirst = True
            End If
            GoTo ContinueLoop
        End If
        If inCodeBlock Then
            If codeFirst Then
                codeLines = ln
                codeFirst = False
            Else
                codeLines = codeLines & vbLf & ln
            End If
            GoTo ContinueLoop
        End If

        ' Markdownテーブル（|セル|セル|）の行かどうか判定する
        If Left(trimmed, 1) = "|" And Right(trimmed, 1) = "|" And Len(trimmed) >= 2 Then
            Set cells = SplitTableCells(trimmed)
            If tablePendingHeader Then
                If IsTableSeparatorRow(cells) Then
                    tablePendingHeader = False
                    Set tableRows = New Collection
                Else
                    ' 直前の行はヘッダーではなく単なるデータ行だった
                    Call FlushTable(outTables, tableHeader, tableRows)
                    Set tableHeader = cells
                    Set tableRows = Nothing
                    tablePendingHeader = True
                End If
                GoTo ContinueLoop
            ElseIf Not tableRows Is Nothing Then
                tableRows.Add cells
                GoTo ContinueLoop
            Else
                Set tableHeader = cells
                tablePendingHeader = True
                GoTo ContinueLoop
            End If
        Else
            If tablePendingHeader Then
                ' ヘッダー候補だけで表として確定しなかった場合は普通の行として扱う
                joined = ""
                For cIdx = 1 To tableHeader.Count
                    If cIdx > 1 Then
                        joined = joined & " / "
                    End If
                    joined = joined & tableHeader(cIdx)
                Next cIdx
                bodyResult = AppendBodyLine(bodyResult, bodyFirst, "・" & joined)
            Else
                Call FlushTable(outTables, tableHeader, tableRows)
            End If
            Set tableHeader = Nothing
            Set tableRows = Nothing
            tablePendingHeader = False
        End If

        ' 箇条書き変換
        If Left(trimmed, 2) = "- " Or Left(trimmed, 2) = "* " Then
            outLine = "・" & Mid(trimmed, 3)
        Else
            outLine = ln
        End If
        bodyResult = AppendBodyLine(bodyResult, bodyFirst, outLine)

ContinueLoop:
    Next i

    Call FlushTable(outTables, tableHeader, tableRows)
    If inCodeBlock And Len(codeLines) > 0 Then
        outCodeBlocks.Add codeLines
    End If

    ' 末尾の空行を除去する
    Do While Right(bodyResult, 1) = Chr(13)
        bodyResult = Left(bodyResult, Len(bodyResult) - 1)
    Loop

    ParseSlideBody = bodyResult

End Function

'=============================================================
' ユーティリティ: 本文文字列に1行追加する（1行目以外はChr(13)区切り）
'=============================================================
Private Function AppendBodyLine(ByVal current As String, ByRef isFirst As Boolean, ByVal newLine As String) As String
    If Not isFirst Then
        current = current & Chr(13)
    End If
    current = current & newLine
    isFirst = False
    AppendBodyLine = current
End Function

'=============================================================
' ユーティリティ: 読み取り中だった表を確定してoutTablesに積む
' （tbl(1)=ヘッダーセルCollection、tbl(2)以降=各行のセルCollection）
'=============================================================
Private Sub FlushTable(ByRef outTables As Collection, ByRef header As Collection, ByRef rows As Collection)
    Dim tbl As Collection
    Dim r As Variant

    If header Is Nothing Then Exit Sub
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub

    Set tbl = New Collection
    tbl.Add header
    For Each r In rows
        tbl.Add r
    Next r
    outTables.Add tbl
End Sub

'=============================================================
' ユーティリティ: Markdownテーブルの1行をセルのCollectionに分解する
'=============================================================
Private Function SplitTableCells(ByVal line As String) As Collection
    Dim inner As String
    Dim parts() As String
    Dim result As Collection
    Dim i As Long

    inner = Trim(line)
    If Left(inner, 1) = "|" Then
        inner = Mid(inner, 2)
    End If
    If Right(inner, 1) = "|" Then
        inner = Left(inner, Len(inner) - 1)
    End If

    parts = Split(inner, "|")

    Set result = New Collection
    For i = 0 To UBound(parts)
        result.Add Trim(parts(i))
    Next i
    Set SplitTableCells = result
End Function

'=============================================================
' ユーティリティ: Markdownテーブルの区切り行（|---|---|等）かどうか判定する
'=============================================================
Private Function IsTableSeparatorRow(ByVal cells As Collection) As Boolean
    Dim c As Variant

    If cells.Count = 0 Then
        IsTableSeparatorRow = False
        Exit Function
    End If
    For Each c In cells
        If Not IsSeparatorCell(CStr(c)) Then
            IsTableSeparatorRow = False
            Exit Function
        End If
    Next c
    IsTableSeparatorRow = True
End Function

Private Function IsSeparatorCell(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim hasDash As Boolean

    s = Trim(s)
    If Len(s) = 0 Then
        IsSeparatorCell = False
        Exit Function
    End If

    hasDash = False
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "-" Then
            hasDash = True
        ElseIf ch = ":" Then
            ' コロンは許容（左寄せ・右寄せ・中央寄せの指定）
        Else
            IsSeparatorCell = False
            Exit Function
        End If
    Next i
    IsSeparatorCell = hasDash
End Function

'=============================================================
' ユーティリティ: プロポーショナルフォントのテキストが折り返されて
' 何行になるか概算する（半角はfontSize*0.55、全角はfontSize分として近似）
'=============================================================
Private Function EstimateWrappedLines(ByVal text As String, ByVal fontSizePt As Single, ByVal widthPt As Single) As Long
    Dim totalWidth As Double
    Dim i As Long
    Dim code As Long
    Dim n As Long

    If Len(text) = 0 Or widthPt <= 0 Then
        EstimateWrappedLines = 1
        Exit Function
    End If

    For i = 1 To Len(text)
        code = AscW(Mid(text, i, 1))
        If code >= 0 And code < 128 Then
            totalWidth = totalWidth + fontSizePt * 0.55
        Else
            totalWidth = totalWidth + fontSizePt
        End If
    Next i

    n = -Int(-(totalWidth / widthPt))  ' 正の値のceiling（VBAにはCeiling関数が無いため）
    If n < 1 Then
        n = 1
    End If
    EstimateWrappedLines = n
End Function

'=============================================================
' ユーティリティ: Chr(13)区切りの本文が必要とする概算の高さ(pt)を返す
'=============================================================
Private Function EstimateTextBlockHeight(ByVal text As String, ByVal fontSizePt As Single, ByVal widthPt As Single) As Single
    Dim lines() As String
    Dim totalLines As Long
    Dim i As Long

    lines = Split(text, Chr(13))

    totalLines = 0
    For i = 0 To UBound(lines)
        totalLines = totalLines + EstimateWrappedLines(lines(i), fontSizePt, widthPt)
    Next i

    EstimateTextBlockHeight = totalLines * fontSizePt * 1.25 + 12
End Function

'=============================================================
' ユーティリティ: 等幅フォント（コードブロック）のテキストが折り返されて
' 何行になるか概算する（半角はfontSize*0.62、全角はfontSize分として近似）
'=============================================================
Private Function EstimateMonoWrappedLines(ByVal text As String, ByVal fontSizePt As Single, ByVal widthPt As Single) As Long
    Dim totalWidth As Double
    Dim i As Long
    Dim code As Long
    Dim n As Long

    If Len(text) = 0 Or widthPt <= 0 Then
        EstimateMonoWrappedLines = 1
        Exit Function
    End If

    For i = 1 To Len(text)
        code = AscW(Mid(text, i, 1))
        If code >= 0 And code < 128 Then
            totalWidth = totalWidth + fontSizePt * 0.62
        Else
            totalWidth = totalWidth + fontSizePt
        End If
    Next i

    n = -Int(-(totalWidth / widthPt))
    If n < 1 Then
        n = 1
    End If
    EstimateMonoWrappedLines = n
End Function

'=============================================================
' ユーティリティ: vbLf区切りのコードブロックが必要とする概算の高さ(pt)を返す
'=============================================================
Private Function EstimateCodeBlockHeight(ByVal text As String, ByVal fontSizePt As Single, ByVal widthPt As Single) As Single
    Dim lines() As String
    Dim totalLines As Long
    Dim i As Long

    lines = Split(text, vbLf)

    totalLines = 0
    For i = 0 To UBound(lines)
        totalLines = totalLines + EstimateMonoWrappedLines(lines(i), fontSizePt, widthPt)
    Next i

    EstimateCodeBlockHeight = totalLines * fontSizePt * 1.15 + 16
End Function

'=============================================================
' 表紙スライド
'=============================================================
Private Sub MakeTitleSlide(prs As Presentation, _
                            title As String, subtitle As String, _
                            author As String, dt As String)

    Dim sld As Slide
    Set sld = prs.Slides.Add(1, ppLayoutBlank)

    ' 背景
    Call SetBgColor(sld, COLOR_LIGHT)

    ' 上部バー
    Call AddRect(sld, 0, 0, sld.Parent.PageSetup.SlideWidth, 14, COLOR_ACCENT)

    ' タイトル背景帯
    Dim bandTop  As Single: bandTop  = CmToPt(4.8)
    Dim bandH    As Single: bandH    = CmToPt(4.5)
    Dim margin   As Single: margin   = CmToPt(2.5)
    Dim bandW    As Single: bandW    = sld.Parent.PageSetup.SlideWidth - margin * 2
    Call AddRect(sld, margin, bandTop, bandW, bandH, COLOR_ACCENT)

    ' タイトルテキスト
    Call AddTextBox(sld, title, margin + 20, bandTop + 15, bandW - 40, _
                    CmToPt(2.2), 36, True, COLOR_WHITE, ppAlignCenter)

    ' サブタイトル
    Call AddTextBox(sld, subtitle, margin + 20, bandTop + CmToPt(2.5), _
                    bandW - 40, CmToPt(1.2), 22, False, COLOR_WHITE, ppAlignCenter)

    ' 作成者・日付
    Call AddTextBox(sld, dt & "　" & author, _
                    margin, CmToPt(14.5), bandW, _
                    CmToPt(0.9), 14, False, COLOR_DARK, ppAlignRight)

    ' 下部バー
    Call AddRect(sld, 0, sld.Parent.PageSetup.SlideHeight - 20, _
                 sld.Parent.PageSetup.SlideWidth, 20, COLOR_ACCENT)

End Sub

'=============================================================
' 目次スライド
'=============================================================
Private Sub MakeTocSlide(prs As Presentation, _
                          titles() As String, cnt As Integer)

    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)

    Call SetBgColor(sld, COLOR_LIGHT)

    Dim hdrH As Single: hdrH = CmToPt(3.2)
    Call AddRect(sld, 0, 0, sld.Parent.PageSetup.SlideWidth, hdrH, COLOR_ACCENT)
    Call AddTextBox(sld, "目次", CmToPt(1.3), CmToPt(0.5), _
                    CmToPt(20), CmToPt(2.0), _
                    28, True, COLOR_WHITE, ppAlignLeft)

    Dim body As String
    Dim i As Integer
    For i = 1 To cnt
        body = body & "  " & i & ".  " & titles(i) & Chr(13)
    Next i

    Call AddTextBox(sld, body, CmToPt(2.5), hdrH + 20, _
                    CmToPt(28), CmToPt(13), _
                    20, False, COLOR_DARK, ppAlignLeft)

    Call AddRect(sld, 0, sld.Parent.PageSetup.SlideHeight - 20, _
                 sld.Parent.PageSetup.SlideWidth, 20, COLOR_ACCENT)

End Sub

'=============================================================
' 本文スライド（本文・表・コードブロックの組み合わせに対応。
' tables・codeBlocksは1スライド分のCollection。それぞれ他の要素と
' 重ならないよう、cursorで縦位置を管理しながら順に積み上げて配置する）
'=============================================================
Private Sub MakeContentSlide(prs As Presentation, title As String, body As String, _
                              tables As Collection, codeBlocks As Collection)

    Dim sld As Slide
    Dim hdrH As Single
    Dim contentLeft As Single
    Dim contentWidth As Single
    Dim cursor As Single
    Dim bottomLimit As Single
    Dim hasExtra As Boolean
    Dim fontSize As Integer
    Dim bodyHeight As Single
    Dim maxBodyHeight As Single
    Dim remaining As Single
    Dim t As Variant
    Dim cb As Variant

    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)

    Call SetBgColor(sld, COLOR_LIGHT)

    hdrH = CmToPt(3.2)
    Call AddRect(sld, 0, 0, sld.Parent.PageSetup.SlideWidth, hdrH, COLOR_ACCENT)
    Call AddTextBox(sld, title, CmToPt(1.3), CmToPt(0.5), _
                    CmToPt(20), CmToPt(2.0), _
                    28, True, COLOR_WHITE, ppAlignLeft)

    contentLeft = CmToPt(2.0)
    contentWidth = CmToPt(29.0)
    cursor = hdrH + 20
    bottomLimit = sld.Parent.PageSetup.SlideHeight - 30

    hasExtra = (tables.Count > 0) Or (codeBlocks.Count > 0)

    If Len(body) > 0 Then
        If hasExtra Then
            fontSize = 18
        Else
            fontSize = 20
        End If

        bodyHeight = EstimateTextBlockHeight(body, fontSize, contentWidth)
        maxBodyHeight = CmToPt(13)
        If bodyHeight > maxBodyHeight Then
            bodyHeight = maxBodyHeight
        End If
        If bodyHeight < 20 Then
            bodyHeight = 20
        End If

        Call AddTextBox(sld, body, contentLeft, cursor, contentWidth, bodyHeight, _
                        fontSize, False, COLOR_DARK, ppAlignLeft)
        cursor = cursor + bodyHeight
    End If

    For Each t In tables
        remaining = bottomLimit - cursor
        If remaining < 20 Then
            remaining = 20
        End If
        cursor = AddTableShape(sld, t, contentLeft, cursor, contentWidth, remaining)
    Next t

    For Each cb In codeBlocks
        remaining = bottomLimit - cursor
        If remaining < 20 Then
            remaining = 20
        End If
        cursor = AddCodeBox(sld, CStr(cb), contentLeft, cursor, contentWidth, remaining)
    Next cb

    Call AddRect(sld, 0, sld.Parent.PageSetup.SlideHeight - 20, _
                 sld.Parent.PageSetup.SlideWidth, 20, COLOR_ACCENT)

End Sub

'=============================================================
' Markdownテーブルを本物のPowerPoint表として挿入する
' （tbl(1)=ヘッダーセルCollection、tbl(2)以降=各行のセルCollection）。
' 行の高さはセルの文字量から概算し、はみ出す場合は最小サイズまで縮小して
' 他の要素との重なりを防ぐ（それでも収まらない場合は手動での内容整理が必要）
'=============================================================
Private Function AddTableShape(sld As Slide, ByVal tbl As Collection, ByVal lft As Single, ByVal tp As Single, _
                                ByVal wd As Single, ByVal maxH As Single) As Single

    Dim header As Collection
    Dim nCols As Long
    Dim nRows As Long
    Dim fontSize As Single
    Dim colWidth As Single
    Dim cellTextWidth As Single
    Dim minRowHeight As Single
    Dim rowHeights() As Single
    Dim r As Long
    Dim totalHeight As Single
    Dim rowCells As Collection
    Dim maxLines As Long
    Dim c As Variant
    Dim ln As Long
    Dim rh As Single
    Dim scaleFactor As Single
    Dim scaledHeight As Single
    Dim shp As Shape
    Dim tbl2 As Table
    Dim cel As Cell
    Dim rowCells2 As Collection
    Dim cellText As String

    Set header = tbl(1)
    nCols = header.Count
    nRows = tbl.Count

    fontSize = 12
    colWidth = wd / nCols
    cellTextWidth = colWidth - 12
    If cellTextWidth < 10 Then
        cellTextWidth = 10
    End If
    minRowHeight = 20

    ReDim rowHeights(1 To nRows)

    totalHeight = 0
    For r = 1 To nRows
        Set rowCells = tbl(r)
        maxLines = 1
        For Each c In rowCells
            ln = EstimateWrappedLines(CStr(c), fontSize, cellTextWidth)
            If ln > maxLines Then
                maxLines = ln
            End If
        Next c
        rh = maxLines * fontSize * 1.3 + 6
        If rh < minRowHeight Then
            rh = minRowHeight
        End If
        rowHeights(r) = rh
        totalHeight = totalHeight + rh
    Next r

    If totalHeight > maxH Then
        scaleFactor = maxH / totalHeight
        totalHeight = 0
        For r = 1 To nRows
            scaledHeight = rowHeights(r) * scaleFactor
            If scaledHeight < minRowHeight * 0.5 Then
                scaledHeight = minRowHeight * 0.5
            End If
            rowHeights(r) = scaledHeight
            totalHeight = totalHeight + scaledHeight
        Next r
    End If

    Set shp = sld.Shapes.AddTable(nRows, nCols, lft, tp, wd, totalHeight)
    Set tbl2 = shp.Table

    For c = 1 To nCols
        tbl2.Columns(c).Width = colWidth
    Next c
    For r = 1 To nRows
        tbl2.Rows(r).Height = rowHeights(r)
    Next r

    For c = 1 To nCols
        Set cel = tbl2.Cell(1, c)
        cel.Shape.TextFrame.TextRange.Text = header(c)
        cel.Shape.TextFrame.TextRange.Font.Size = fontSize
        cel.Shape.TextFrame.TextRange.Font.Bold = True
        cel.Shape.TextFrame.TextRange.Font.Color.RGB = COLOR_WHITE
        cel.Shape.TextFrame.TextRange.Font.Name = "メイリオ"
        cel.Shape.Fill.Solid
        cel.Shape.Fill.ForeColor.RGB = COLOR_ACCENT
    Next c

    For r = 2 To nRows
        Set rowCells2 = tbl(r)
        For c = 1 To nCols
            If c <= rowCells2.Count Then
                cellText = rowCells2(c)
            Else
                cellText = ""
            End If
            Set cel = tbl2.Cell(r, c)
            cel.Shape.TextFrame.TextRange.Text = cellText
            cel.Shape.TextFrame.TextRange.Font.Size = fontSize
            cel.Shape.TextFrame.TextRange.Font.Color.RGB = COLOR_DARK
            cel.Shape.TextFrame.TextRange.Font.Name = "メイリオ"
            cel.Shape.Fill.Solid
            cel.Shape.Fill.ForeColor.RGB = COLOR_WHITE
        Next c
    Next r

    AddTableShape = tp + totalHeight + 15

End Function

'=============================================================
' コードブロック（```～```で囲まれた部分）を等幅フォント（Courier New）の
' テキストボックスとして表示する（ディレクトリツリー等のファイル構成表示向け）。
' 収まらない場合はフォントを段階的に縮小し、他の要素との重なりを防ぐ
'=============================================================
Private Function AddCodeBox(sld As Slide, codeText As String, ByVal lft As Single, ByVal tp As Single, _
                             ByVal wd As Single, ByVal maxH As Single) As Single

    Dim fontSize As Single
    Dim innerWidth As Single
    Dim ht As Single
    Dim trySize As Single
    Dim candidate As Single
    Dim shp As Shape

    fontSize = 13
    innerWidth = wd - 20

    ht = EstimateCodeBlockHeight(codeText, fontSize, innerWidth)

    If ht > maxH Then
        For trySize = fontSize - 1 To 8 Step -1
            candidate = EstimateCodeBlockHeight(codeText, trySize, innerWidth)
            If candidate <= maxH Then
                fontSize = trySize
                ht = candidate
                Exit For
            End If
        Next trySize
        If ht > maxH Then
            ht = maxH
        End If
    End If

    Call AddRect(sld, lft, tp, wd, ht, COLOR_CODE_BG)
    Call AddRect(sld, lft, tp, 4, ht, COLOR_ACCENT)

    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, lft + 12, tp + 8, innerWidth - 8, ht - 16)
    With shp.TextFrame
        .WordWrap = msoTrue
        .AutoSize = ppAutoSizeNone
        With .TextRange
            .Text = Replace(codeText, vbLf, Chr(13))
            .Font.Size = fontSize
            .Font.Color.RGB = COLOR_DARK
            .Font.Name = "Courier New"
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
    End With

    AddCodeBox = tp + ht + 15

End Function

'=============================================================
' まとめスライド（背景を濃紺で強調）
'=============================================================
Private Sub MakeSummarySlide(prs As Presentation, title As String, body As String)

    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)

    Call SetBgColor(sld, COLOR_ACCENT)

    Call AddTextBox(sld, title, CmToPt(1.3), CmToPt(0.7), _
                    CmToPt(30), CmToPt(2.3), _
                    32, True, COLOR_WHITE, ppAlignCenter)

    ' 区切り線
    Call AddRect(sld, CmToPt(1.3), CmToPt(3.0), _
                 CmToPt(31.2), 2, COLOR_WHITE)

    Call AddTextBox(sld, body, CmToPt(2.0), CmToPt(3.5), _
                    CmToPt(29), CmToPt(13), _
                    20, False, COLOR_WHITE, ppAlignLeft)

End Sub

'=============================================================
' ユーティリティ: 背景色設定
'=============================================================
Private Sub SetBgColor(sld As Slide, col As Long)
    sld.Background.Fill.Solid
    sld.Background.Fill.ForeColor.RGB = col
End Sub

'=============================================================
' ユーティリティ: 矩形追加
'=============================================================
Private Sub AddRect(sld As Slide, l As Single, t As Single, _
                    w As Single, h As Single, col As Long)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, l, t, w, h)
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = col
    shp.Line.Visible = msoFalse
End Sub

'=============================================================
' ユーティリティ: テキストボックス追加
'=============================================================
Private Sub AddTextBox(sld As Slide, txt As String, _
                        l As Single, t As Single, w As Single, h As Single, _
                        fontSize As Integer, isBold As Boolean, _
                        col As Long, align As PpParagraphAlignment)
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)

    With shp.TextFrame
        .WordWrap = msoTrue
        .AutoSize = ppAutoSizeNone
        With .TextRange
            .Text = txt
            .Font.Size  = fontSize
            .Font.Bold  = isBold
            .Font.Color.RGB = col
            .Font.Name  = "メイリオ"
            .ParagraphFormat.Alignment = align
        End With
    End With
End Sub
