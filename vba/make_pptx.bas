Attribute VB_Name = "MakePPTX"
'=============================================================
' PowerPoint定型スライド自動生成マクロ
' 使い方: PowerPoint VBAエディタ(Alt+F11)からMakePPTXを実行
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
        .SlideWidth  = CentimetersToPoints(33.867)
        .SlideHeight = CentimetersToPoints(19.05)
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
    Dim bandTop  As Single: bandTop  = CentimetersToPoints(4.8)
    Dim bandH    As Single: bandH    = CentimetersToPoints(4.5)
    Dim margin   As Single: margin   = CentimetersToPoints(2.5)
    Dim bandW    As Single: bandW    = sld.Parent.PageSetup.SlideWidth - margin * 2
    Call AddRect(sld, margin, bandTop, bandW, bandH, COLOR_ACCENT)

    ' タイトルテキスト
    Call AddTextBox(sld, title, margin + 20, bandTop + 15, bandW - 40, _
                    CentimetersToPoints(2.2), 36, True, COLOR_WHITE, ppAlignCenter)

    ' サブタイトル
    Call AddTextBox(sld, subtitle, margin + 20, bandTop + CentimetersToPoints(2.5), _
                    bandW - 40, CentimetersToPoints(1.2), 22, False, COLOR_WHITE, ppAlignCenter)

    ' 作成者・日付
    Call AddTextBox(sld, dt & "　" & author, _
                    margin, CentimetersToPoints(14.5), bandW, _
                    CentimetersToPoints(0.9), 14, False, COLOR_DARK, ppAlignRight)

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

    Dim hdrH As Single: hdrH = CentimetersToPoints(3.2)
    Call AddRect(sld, 0, 0, sld.Parent.PageSetup.SlideWidth, hdrH, COLOR_ACCENT)
    Call AddTextBox(sld, "目次", CentimetersToPoints(1.3), CentimetersToPoints(0.5), _
                    CentimetersToPoints(20), CentimetersToPoints(2.0), _
                    28, True, COLOR_WHITE, ppAlignLeft)

    Dim body As String
    Dim i As Integer
    For i = 1 To cnt
        body = body & "  " & i & ".  " & titles(i) & Chr(13)
    Next i

    Call AddTextBox(sld, body, CentimetersToPoints(2.5), hdrH + 20, _
                    CentimetersToPoints(28), CentimetersToPoints(13), _
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

    Dim hdrH As Single: hdrH = CentimetersToPoints(3.2)
    Call AddRect(sld, 0, 0, sld.Parent.PageSetup.SlideWidth, hdrH, COLOR_ACCENT)
    Call AddTextBox(sld, title, CentimetersToPoints(1.3), CentimetersToPoints(0.5), _
                    CentimetersToPoints(20), CentimetersToPoints(2.0), _
                    28, True, COLOR_WHITE, ppAlignLeft)

    Call AddTextBox(sld, body, CentimetersToPoints(2.0), hdrH + 20, _
                    CentimetersToPoints(29), CentimetersToPoints(13), _
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

    Call AddTextBox(sld, title, CentimetersToPoints(1.3), CentimetersToPoints(0.7), _
                    CentimetersToPoints(30), CentimetersToPoints(2.3), _
                    32, True, COLOR_WHITE, ppAlignCenter)

    ' 区切り線
    Call AddRect(sld, CentimetersToPoints(1.3), CentimetersToPoints(3.0), _
                 CentimetersToPoints(31.2), 2, COLOR_WHITE)

    Call AddTextBox(sld, body, CentimetersToPoints(2.0), CentimetersToPoints(3.5), _
                    CentimetersToPoints(29), CentimetersToPoints(13), _
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
