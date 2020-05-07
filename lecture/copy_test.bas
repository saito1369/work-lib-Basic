Attribute VB_Name = "AutoPowerPointDecoration"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub collectPpt()

  Dim   Fname As String
  Fname = "C:\Users\saito\Desktop\a.pptx"

  ' 今開いているスライド
  Dim pTo  As Presentation
  Set PTo = Application.ActivePresentation

  ' 保存する資料のファイル名
  Dim sPptFileName As String
  sPptFileName="test_intg.pptm"

  ' ここからはこのファイルが Active となるので注意する.
  'ActivePresentation.SaveAs (sPptFileName)

  ' ppt ファイルの取り込み
  Dim pFr As Presentation ' ppt ファイルを開きます.
  Set pFr = Presentations.Open(FileName:=Fname,ReadOnly:=msoFalse)

  Call copySlide_Fmt(pFr,pTo,",")
  ActivePresentation.SaveAs (sPptFileName)

End Sub

' slide のスタイルを元の slide の状態に保ったままコピーします
Private Sub copySlide_Fmt(pFr As Presentation, pTo As Presentation, delm As String)

  ' copy するページ
  Dim Pages As Variant
  Pages = Array(1,2,3,4,5,6,7,8,9,10)

  ' keep formatting copy
  ' http://img2.tapuz.co.il/forums/1_88724584.txt
  Dim q  As Integer
  For q = 0 To UBound(Pages) ' 1 page ずつコピーする
    Call ClearClipboard()
    Dim sFr As Slide
    Set sFr = pFr.Slides(Pages(q))
    sFr.Copy
    ' copy と paste の間に休みを入れる(始まり)
    DoEvents
    Sleep 10 ' 10(ms) の休憩
    DoEvents
    ' copy と paste の間に休みを入れる(終わり)
    With pTo.Slides.Paste
      .Design      = sFr.Design
      .ColorScheme = sFr.ColorScheme
      If sFr.FollowMasterBackground = False Then
            .FollowMasterBackground = False
        With .Background.Fill
          .Visible   = sFr.Background.Fill.Visible
          .ForeColor = sFr.Background.Fill.ForeColor
          .BackColor = sFr.Background.Fill.BackColor
        End With
        Select Case sFr.Background.Fill.Type
          Case Is = msoFillTextured
            Select Case sFr.Background.Fill.TextureType
              Case Is = msoTextruePreset
                .Background.Fill.PresetTextured _
                           sFr.Background.Fill.PresetTexture
              Case Is = msoTextureUserDefined
            End Select
          Case Is = msoFillSolid
            .Background.Fill.Transparency = 0#
            .Background.Fill.Solid
          Case Is = msoFillPicture
            With sFr
              If .Shapes.Count > 0 Then .Shapes.Range.Visible = False
              bMasterShapes = .DisplayMasterShapes
              .DisplayMasterShapes = False
              .Export pFr.Path & .SlideID & ".png", "PNG"
            End With
            .Background.Fill.UserPicture _
                       pFr.Path & sFr.SlideID & ".png"
            Kill (pFr.Path & sFr.SlideID & ".png")
            With sFr
              .DisplayMasterShapes = bMasterShapes
              If .Shapes.Count > 0 Then .Shapes.Range.Visible = True
            End With
          Case Is = msoFillPatterned
            .Background.Fill.Patterned _
                       (sFr.Background.Fill.Pattern)
          Case Is = msoFillGradient
            Select Case sFr.Background.Fill.GradientColorType
              Case Is = msoGradientTwoColors
                .Background.Fill.TwoColorGradient _
                           sFr.Background.Fill.GradientStyle, _
                           sFr.Background.Fill.GradientVariant
              Case Is = msoGradientPresetColors
                .Background.Fill.PresetGradient _
                           sFr.Background.Fill.GradientStyle, _
                           sFr.Background.Fill.GradientVariant, _
                           sFr.Background.Fill.PresetGradientType
              Case Is = mstGradientOneColor
                .Background.Fill.OneColorGradient _
                           sFr.Background.Fill.GradientStyle, _
                           sFr.Background.Fill.GradientVariant, _
                           sFr.Background.Fill.GradientDegree
            End Select
          Case Is = msoFillBackground
            ' Only applicate to shapes.
        End Select
      End If
    End With
  Next q
  pFr.Close

End Sub

