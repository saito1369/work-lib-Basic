Attribute VB_Name = "TiledFig"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub addPictures()

  ' 既存の powerpoint 資料に図を埋め込む
  ' 設定ファイルを読み込み
  ' powerpoint のページ数\t図のファイル名\ttop位置(cm)\tleft 位置(cm)\t倍率\n
  ' 1    test.png    0.0     2.5    2.5
  ' 1    test2.png   12.0    2.5    2.5
  '

  Dim delm As String
  Dim os   As Integer
  os = 0
  delm = "\"
  If Application.OperatingSystem Like "Macintosh*" Then
    delm = ":"
    os   = 1   ' flag for mac
  End If

  ' text file
  Dim sCurrentFolder As String   ' ppt がある folder
  Dim sNotesFileName As String   ' text file 名
  sCurrentFolder = ActivePresentation.Path & delm
  sNotesFileName = Mid$(ActivePresentation.Name, 1, InStr(ActivePresentation.Name, ".") - 1)
  sNotesFileName = sNotesFileName & ".txt"
  If MsgBox("do you want to make the table of contents using the file '" & sNotesFileName & "'?", vbYesNo) = vbNo Then
    sNotesFileName = InputBox("input file name", "図挿入設定ファイル名の入力", sNotesFileName)
  End If

  ' is it there? quit if not
  If Len(Dir$(sCurrentFolder & sNotesFileName)) = 0 Then
    MsgBox (sCurrentFolder & sNotesFileName & " is missing")
    Exit Sub
  End If

  ' open the file and go to work
  Dim iNotesFileNum As Integer
  iNotesFileNum = FreeFile()
  Open sCurrentFolder & sNotesFileName For Input As iNotesFileNum

  Dim Pages()  As Integer ' powerpoint のページ数(1-based)
  Dim Fnames() As String  ' png file name
  Dim lefts()  As Single  ' left
  Dim tops()   As Single  ' top
  Dim ratios() As Single  ' ratio

  ' file open
  Dim sBuf   As String
  Dim aBuf() As String
  Dim fnum   As Long
  fnum = 0
  Do Until EOF(iNotesFileNum)
    Line Input #iNotesFileNum, sBuf
    If left(sBuf,1) = "#" Then    '# で始まっている行はコメント
      ' do nothing
    ElseIf Len(sBuf) > 0 Then    ' 空行でなければ読み込み
      ReDim Preserve Pages(fnum)
      ReDim Preserve Fnames(fnum)
      ReDim Preserve lefts(fnum)
      ReDim Preserve tops(fnum)
      ReDim Preserve ratios(fnum)

      aBuf = Split(sBuf,vbTab)

      Pages(fnum)  = aBuf(0)
      lefts(fnum)  = aBuf(2)
      tops(fnum)   = aBuf(3)
      ratios(fnum) = aBuf(4)

      Dim Fname As String
      Fname = aBuf(1)
      If os = 1 Then  'mac
        Fname = Replace(Fname,"\",delm)
        Fname = Replace(Fname,"/",delm)
        ' ファイル名だけ書いてある場合は, カレントディレクトリだと仮定する.
        If InStr(Fname, delm) = 0 Then Fname = sCurrentFolder & Fname
      Else 'win
        ' 基本的には, フォルダ区切りは windows 形式で書くことにする.
        'Fname = Replace(Fname, ":",delm)
        'If InStr(Fname, delm) = 0 Then Fname = sCurrentFolder & Fname
        ' これだとうまくいった. Why ? 2015/06/30
        Fname = sCurrentFolder & Fname
      End If
      Fnames(fnum)=Fname
      fnum = fnum + 1
    End If
  Loop
  Close iNotesFileNum

  Dim Pptt As Presentation
  Set Pptt = Application.ActivePresentation

  Dim k As Long
  For k = 0 To UBound(Pages)
    Dim ShapeObj As Shape
    Dim pleft    As Single
    Dim ptop     As Single
    pleft = lefts(k)/ 0.3528 * 10.0  ' cm -> points
    ptop  = tops(k) / 0.3528 * 10.0

    ' slide が無ければ追加する 2015/04/16
    Dim SlideObj As Slide
    If Pages(k) <= Pptt.Slides.Count Then  ' Pptt.Slides.Count = スライドの枚数
      Set SlideObj = Pptt.Slides(Pages(k)) ' Slide が既にある場合
    Else
      Set SlideObj = Pptt.Slides.Add(Pages(k),ppLayoutBlank) ' Slide を新たに追加する場合
    End If
    Set ShapeObj = SlideObj.Shapes.addPicture(Fnames(k),msoFalse,msoTrue,pleft,ptop,-1,-1)
    With ShapeObj
      .ScaleHeight ratios(k), msoTrue
      .ScaleWidth  ratios(k), msoTrue
      .ZOrder      msoSendToBack        ' 最背面に移動
      .Name = "addedpicture"
    End With
  Next k

End Sub

Sub removeAddedPictures()
  Dim SlideObj   As Slide
  Dim ShapeObj   As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      'If ShapeObj.Type = msoGroup Then            ' グループ
        If ShapeObj.Name = "addedpicture" Then      ' 名前が "addedpicture"
          ShapeObj.Delete
        End If
      'End If
    Next ShapeIndex
  Next SlideObj
End Sub
