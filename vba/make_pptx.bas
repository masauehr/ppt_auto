Attribute VB_Name = "MakePPTX"
'=============================================================
' PowerPoint定型スライド自動生成マクロ
'
' 収録マクロ:
'   MakePPTX             - コード内に直書きした雛形からスライドを生成
'   MakePPTXFromMarkdown - Markdownファイルを読み込んでスライドを生成
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
Private Const COLOR_ACCENT As Long = &H7D491F  ' BGRで格納（VBA仕様）→ 実際は #1F497D
Private Const COLOR_WHITE  As Long = &HFFFFFF
Private Const COLOR_DARK   As Long = &H262626
Private Const COLOR_LIGHT  As Long = &HF2F2F2

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
            Call MakeContentSlide(prs, titles(i), bodies(i))
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

    Dim slideTitles() As String
    Dim slideBodies() As String
    Dim slideTypes() As String
    Dim slideCount As Integer
    Call ParseSlides(bodyText, slideTitles, slideBodies, slideTypes, slideCount)

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
            Call MakeContentSlide(prs, slideTitles(i), slideBodies(i))
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
' ユーティリティ: "## " 見出しでスライドに分割する
'=============================================================
Private Sub ParseSlides(ByVal src As String, _
                         ByRef outTitles() As String, ByRef outBodies() As String, _
                         ByRef outTypes() As String, ByRef outCount As Integer)

    Dim lines() As String
    lines = Split(src, vbLf)

    Dim maxSlides As Integer
    maxSlides = UBound(lines) + 1
    If maxSlides < 1 Then maxSlides = 1

    ReDim outTitles(1 To maxSlides)
    ReDim outBodies(1 To maxSlides)
    ReDim outTypes(1 To maxSlides)
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
                outBodies(outCount) = FormatBody(curBodyLines)
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
        outBodies(outCount) = FormatBody(curBodyLines)
        outTypes(outCount) = curType
    End If

End Sub

'=============================================================
' ユーティリティ: スライド本文を整形する
' （"- "箇条書きを"・"に変換し、Chr(13)区切りの1文字列にする）
'=============================================================
Private Function FormatBody(ByVal raw As String) As String

    Dim lines() As String
    lines = Split(raw, vbLf)

    Dim result As String
    Dim first As Boolean
    first = True

    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = lines(i)

        Dim trimmed As String
        trimmed = Trim(ln)
        If Left(trimmed, 2) = "- " Or Left(trimmed, 2) = "* " Then
            ln = "・" & Mid(trimmed, 3)
        End If

        If Not first Then
            result = result & Chr(13)
        End If
        result = result & ln
        first = False
    Next i

    ' 末尾の空行を除去する
    Do While Right(result, 1) = Chr(13)
        result = Left(result, Len(result) - 1)
    Loop

    FormatBody = result

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
' 本文スライド
'=============================================================
Private Sub MakeContentSlide(prs As Presentation, title As String, body As String)

    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)

    Call SetBgColor(sld, COLOR_LIGHT)

    Dim hdrH As Single: hdrH = CmToPt(3.2)
    Call AddRect(sld, 0, 0, sld.Parent.PageSetup.SlideWidth, hdrH, COLOR_ACCENT)
    Call AddTextBox(sld, title, CmToPt(1.3), CmToPt(0.5), _
                    CmToPt(20), CmToPt(2.0), _
                    28, True, COLOR_WHITE, ppAlignLeft)

    Call AddTextBox(sld, body, CmToPt(2.0), hdrH + 20, _
                    CmToPt(29), CmToPt(13), _
                    20, False, COLOR_DARK, ppAlignLeft)

    Call AddRect(sld, 0, sld.Parent.PageSetup.SlideHeight - 20, _
                 sld.Parent.PageSetup.SlideWidth, 20, COLOR_ACCENT)

End Sub

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
