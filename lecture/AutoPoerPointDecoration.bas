Attribute VB_Name = "AutoPowerPointDecoration"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' text file を読み込む format の変更
' see Private Sub fireRead2
' 1. 情報目印最初と最後に "::" を入れる
' 2. 目印本体は 3 文字の文字列(何でも良い) 例: @@@
' 3. hBufs("@@@") で文字列が取得できる
' 4. 得られた文字列を parse して知りたい情報を得る
'
' ::@@@::
' (何か書く)
' ::$$$::
' (何か書く)

' bufRead_TOC:    |||
' bufRead_File:   +++
' bufRead_pList:  ###,^^^
' bufRead_nList:  $$$,&&&,@@@,===
' bufRead_exList: %%%
' bufRead_siz:   !!!,???

' 2015/01/06
' Scripting.Dictionary for MacOSX
'
Private Const C_EXC       As Integer = -9999
Private Const C_GRAYTOC   As String  = "grayTOC"
Private Const C_PAGENUM   As String  = "pagenum"
Private Const C_COUNTBOX  As String  = "countbox"
Private Const C_CONTENT   As String  = "tcontents"
Private Const C_HIERARCHY As String  = "thierarchy"
Private Const C_CMM       As String  = ","
Private Const C_TBB       As String  = vbTab
Private Const C_TITLE     As String  = "目次"
Private Const C_PLCHOLDER As String  = "Placeholder"
Private Const C_TITLE1    As String  = "Title"
Private Const C_FSIZE     As Integer = 8
Private Const C_XPOSR     As Integer = 610    ' pageListContents default X 座標
Private Const C_YPOSR     As Integer = 5      ' pageListContents
Private Const C_WIDER     As Integer = 140    ' pageListContents
Private Const C_XPOSL     As Integer = 5      ' pageListHierarchy
Private Const C_YPOSL     As Integer = 1      ' pageListHierarchy
Private Const C_WIDEL     As Integer = 200    ' pageListHierarchy
Private Const C_PSIZE     As Integer = 10     ' pageNum
Private Const C_XPOSB     As Integer = 686    ' pageNum
Private Const C_YPOSB     As Integer = 527    ' pageNum
Private Const C_HEIGB     As Integer = 29     ' pageNum
Private Const C_WIDEB     As Integer = 40     ' pageNum
Private Const C_FSIZE_TOC As Integer = 20     ' resizeFontSizeOfGrayTOC
Private Const C_TSIZE_TOC As Integer = 36     ' resizeFontSIzeOfGrayTOC
Private Const C_XPOSS     As Integer = 5      ' source_copy
Private Const C_YPOSS     As Integer = 5      ' source_copy
Private Const C_XWIDS     As Integer = 700    ' source_copy
Private Const C_YWIDS     As Integer = 600    ' source_copy
Private Const C_SSIZE     As Integer = 12     ' source_copy
Private Const C_MAXH      As Integer = 2      ' zmake_TOC_for_ListContents_from_hier


Private Const COMT_MARK   As String  = "'"    ' comment 行の指定
Private Const INFO_MARK   As String  = "::"   ' 情報入力の目印
Private Const GTOC_MARK   As String  = "|||"  ' grayTOC 用の目印           grayTOC
Private Const CPPT_MARK   As String  = "+++"  ' collectPpt 用の目印        collectPpt
Private Const SKIP_MARK   As String  = "---"  ' 無視するページ             skipSlide
Private Const TOCS_MARK   As String  = "###"  ' 目次とページ数             pageListContents, pageListHierarchy
Private Const NOWT_MARK   As String  = "$$$"  ' 目次を書きださないページ   pageListContents
Private Const IFNT_MARK   As String  = "!!!"  ' 書きだす font 情報         pageListContents
Private Const TOHI_MARK   As String  = "^^^"  ' 目次とページ数             pageListHierarchy
Private Const NOHI_MARK   As String  = "&&&"  ' 目次を書きださないページ   pageListHierarchy
Private Const IFHI_MARK   As String  = "???"  ' 書きだす font 情報         pageListHierarchy
Private Const FLIP_MARK   As String  = "%%%"  ' 置換文字列リスト           zFlipTextForStudent
Private Const DELS_MARK   As String  = "@@@"  ' 削除スライド               zFlipTextForStudent
Private Const ALIV_MARK   As String  = "==="  ' 保存スライド(削除 @@@優先) zFlipTextForStudent
Private Const SOUR_MARK   As String  = "sss"  ' ファイルの中身を直接 ppt に書きだす source_copy

Private Const C_FNAME     As String  = "Arial"
Private Const C_ENAME     As String  ="ＭＳ Ｐゴシック"
Private Const C_SNAME     As String  ="ＭＳ ゴシック"    ' source_copy

Private Const C_LWID      As Integer = 5      ' watermarkSlide 線の太さ

'Private Const BLACK = RGB(0,0,0)
'Private Const GRAY  = RGB(150,150,150)
'Private Const RED   = RGB(255,0,0)
'Private Const GRAY2 = RGB(70,70,70)
'Private Const WHITE = RGB(255,255,255)
'Private Const GREEN = RGB(0,150,0)

Public Sub pv_000_standard()
  Call pv_0000_collectPpt()
  Call pv_0010_grayTOC()
  Call pv_0011_pageCountsBox2()
  Call pv_0012_pageNum2()
  Call pv_0013_pageListContents2()
  Call pv_0014_pageListHierarchy2()
End Sub


Public Sub pv_0010_grayTOC()
  ' 階層目次の指示が書かれた設定ファイル(text file)を読み込み,
  ' 階層目次(灰色)を然るべきページに追加して別ファイルとして保存する.
  ' 追加するとページ数がずれるので, ずれを補正して保存した別ファイルの設定ファイルとして保存する.
  ' 例: hoge.ppt に階層的目次を追加する. 設定ファイルは hoge.txt
  ' hoge.txt の中身:
  ' ::|||::
  ' 本日の内容
  ' *  1. 背景              1
  ' ** 1.1. はじめに        2
  ' ** 1.2. 歴史            3
  ' *  2. 目的              4
  ' hoge.txt の format:
  ' ::|||::
  ' *   (\t)(階層1の目次)(\t)(\t)(書きだすページ数)　注: このページの前に追加する. このページ以降のページ数がずれる.
  ' **  (\t)(階層2の目次)(\t)(\t)(書きだすページ数)
  ' *** (\t)(階層3の目次)(\t)(\t)(書きだすページ数)
  '
  ' output:
  ' 階層目次が追加された ppt file: hoge_grayTOC.ppt
  ' 追加階層目次のページだけずれを補正, 目次をカウントしないスライドとして追加した txt file: hoge_grayTOC.txt
  '

  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)

  ' iNc***(1 から hire, 1 から page 数) の二次元配列
  ' iNcIdx(hier,page) = このページの前では, この id の Toc だけ黒く書きます(他は灰色)
  ' iNcTtl(hier,page) = このページの前に, このタイトルで目次を書きます.
  ' iNcToc(hier,page) = このページの前に書く目次部分(改行で split すると配列になる)
  ' iNcOff(hier,page) = このページの前に置く目次ページの数 (offset を計算するときに使う)
  Dim iNcIdx()  As String
  Dim iNcTtl()  As String
  Dim iNcToc()  As String
  Dim iNcOff()  As Integer
  Call bufRead_TOC(hBufs(GTOC_MARK),iNcIdx(),iNcTtl(),iNcToc(),iNcOff(),C_TBB)

  Dim sPptFileName As String
  sPptFileName=newFileName("_gray.pptm","do you want to save the file As ",1)
  ' 別名で上書き保存
  ' ここからはこのファイルが Active となるので注意する.
  If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

  Call add_grayTOC(iNcIdx(),iNcTtl(),iNcToc(),iNcOff())

  ' page のズレを計算する
  ' 新しいページ数 = 古いページ数 + offset(古いページ数)
  Dim offset() As Integer
  Call offset_grayTOC(offset(), iNcOff())

  Dim sGrayFilePath As String
  Dim sGrayFileName As String
  sGrayFileName=newFileName(".txt","do you want to save the file As ",1)
  sGrayFilePath = sCurrentFolder & sGrayFileName
  'Call offsetPrint(sGrayFilePath,hBufs,offset(),Nothing)
  Call offsetPrint(sGrayFilePath,hBufs,offset()) '2014/04/06

  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

End Sub

Public Sub removeGrayTOC()
  Dim npage As Integer
  npage = ActivePresentation.Slides.Count
  Dim i As Integer
  For i = npage To 1 Step -1
    'ActivePresentation.Slides(i).Select
    If InStr(ActivePresentation.Slides(i).Name,C_GRAYTOC) Then
      ActivePresentation.Slides(i).Delete
    End If
  Next i
End Sub

Public Sub resizeFontSizeOfGrayTOC()
  'Dim Blk
  'Blk=RGB(0,0,0)
  Dim Fsize As Integer
  Dim Tsize As Integer
  
  Fsize = C_FSIZE_TOC
  Tsize = C_TSIZE_TOC
  
  If MsgBox("font resize for gray TOC to " & Fsize & " ?",vbYesNo) = vbNo Then
    Fsize = InputBox("Input new font size for gray TOC",Fsize)
  End If
  If MsgBox("font resize for title TOC to " & Tsize & " ?",vbYesNo) = vbNo Then
    Tsize = InputBox("Input new font size for title TOC",Tsize)
  End If

  Dim npage As Integer
  npage = ActivePresentation.Slides.Count
  Dim i      As Integer
  Dim oShape As Shape
  For i = npage To 1 Step -1
    If InStr(ActivePresentation.Slides(i).Name,C_GRAYTOC) Then
      For Each oShape In ActivePresentation.Slides(i).Shapes
        'MsgBox("page=" & i & "  shape type = " & oShape.Type)
        'MsgBox("msoPlaceholder=" & msoPlaceholder & " name = " & oShape.Name)
        If InStr(oShape.Name,C_TITLE1) Then     ' 2014/10/08 oShape.Name が "Title" を含んでたら変更
          oShape.TextFrame2.TextRange.Characters.Font.Size=Tsize
        End If
        If InStr(oShape.Name,C_PLCHOLDER) Then
          ' 適当にやったらできた. これでいいのか?
          oShape.TextFrame2.TextRange.Characters.Font.Size=Fsize
          'oShape.TextFrame.TextRange.Characters.Font.Color.RGB=Blk
        End If
      Next oShape
    End If
  Next i
End Sub

Public Sub pv_1000_removeGrayTOC_asNewName() 'for print (pdf) hand out
  Dim npage As Integer
  npage = ActivePresentation.Slides.Count
  Dim i As Integer
  For i = npage To 1 Step -1
    'ActivePresentation.Slides(i).Select
    If InStr(ActivePresentation.Slides(i).Name,C_GRAYTOC) Then
      ActivePresentation.Slides(i).Delete
    End If
  Next i
  
  Dim sNotesFilePath As String ' dummy
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",0,0)
  Dim sPptFileName As String
  spptFileName = newFileName("_delTOC.pptm","do you want to save the file as ",1)
  If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

End Sub

' hBufs を offset にもとづいて補正した hash string (nBufs) を返します.
Private Sub offsetHash(hBufs, nBufs, offset() As Integer, Optional pExist As Variant)

  Dim keys() As Variant
  keys = hBufs.Keys
  Dim key As Variant

  For Each key In keys
    Dim sBuf As String
    If key = GTOC_MARK Then
      sBuf=""  ' 何でかよくわからない. scope が効いてないのか ?
      Call bufWrit_TOC(sBuf,hBufs(key),C_TBB,offset(),pExist)
    ElseIf key = CPPT_MARK Then
      sBuf=""
      Call bufWrit_File(sBuf,hBufs(key),C_TBB,offset(),pExist)
    ElseIf key = TOCS_MARK Or key = TOHI_MARK Then
      sBuf=""
      Call bufWrit_pList(sBuf,hBufs(key),C_TBB,offset(),pExist)
    ElseIf key = NOWT_MARK Or key = NOHI_MARK Or key =DELS_MARK Or key = ALIV_MARK Or key = SKIP_MARK Then
      sBuf=""
      Call bufWrit_nList(sBuf,hBufs(key),C_CMM,offset(),pExist)
    ElseIf key = FLIP_MARK Then
      sBuf=""
      Call bufWrit_exList(sBuf,hBufs(key),C_TBB,offset(),pExist)
    ElseIf key = IFNT_MARK Or key = IFHI_MARK Then
      sBuf=""
      Call bufWrit_siz(sBuf,hBufs(key))
    End If
    If nBufs.Exists(key) Then
      nBufs(key) = nBufs(key) & vbNewLine & sBuf
    Else
      nBufs.Add key,sBuf
    End If
  Next key

End Sub

' offset 補正した text file を書き出します.
Private Sub offsetPrint(sFilePath As String, hBufs, offset() As Integer, Optional pExist As Variant)
  Set nBufs   = New Dictionary ' 読み込んだ文字列
  Call offsetHash(hBufs,nBufs,offset(),pExist)
  Call hashPrint(nBufs,sFilePath)
End Sub

' hash で得られた String を決まった format で書き出します
Private Sub hashPrint(hash, sFilePath As String)
  Dim keys() As Variant
  keys = hash.Keys
  Dim key As Variant

  Dim prnt As String
  For Each key In keys
    prnt = prnt & INFO_MARK & key & INFO_MARK & vbNewLine
    prnt = prnt & hash(key)         & vbNewLine
    'MsgBox("key= " & key & vbTab & "value = " & hash(key))
  Next key

  Dim iFileNum As Integer
  iFileNum = FreeFile()
  Open sFilePath For Output As #iFileNum
  Print #iFileNum, prnt
  Close #iFileNum
End Sub

' gray 目次を挿入したことによる offset 値を計算します.
Private Sub offset_grayTOC(offset() As Integer, iNcOff() As Integer)
  Dim nhier As Integer
  nhier = UBound(iNcOff,1)
  Dim npage As Integer
  npage = UBound(iNcOff,2)
  Dim off() As Integer
  ReDim off(npage)
  Dim hier As Integer
  For hier = 1 To nhier
    Dim page As Integer
    For page = 1 To npage
      off(page) = off(page) + iNcOff(hier,page)
    Next page
  Next hier

  ReDim Preserve offset(npage)
  Dim diff As Integer
  diff = 0
  For page = 1 To npage
    offset(page)= off(page) + diff
    diff = diff + off(page)
  Next page
End Sub

' text file の情報に基づいて gray 目次スライドを挿入します.
Private Sub add_grayTOC(iNcIdx() As String, iNcTtl() As String, iNcToc() As String, iNcOff() As Integer)
  Dim nhier As Integer
  nhier = UBound(iNcTtl,1) ' 階層の数
  Dim npage As Integer
  npage = ActiveWindow.Presentation.Slides.Count ' スライドの枚数

  If npage > UBound(iNcTtl,2) Then
    npage  = UBound(iNcTtl,2)
  End If

  ' 色設定
  Dim Blk, Gry
  'Blk=BLACK
  'Gry=GRAY
  Blk = RGB(0, 0, 0)
  Gry = RGB(150, 150, 150)

  Dim offs As Integer
  offs = 0
  Dim page As Integer
  For page = 1 To npage
    Dim hier As Integer
    For hier = 1 To nhier
      Dim title As String
      title = iNcTtl(hier,page)
      ' スライド追加と灰色メソッドで階層目次を自動書き出し
      If Not IsNull(title) And Not title = "" Then        ' 目次タイトルが存在すれば
        If page <= npage Then ' ページがあれば
          Dim pageo As Integer
          pageo = page + offs
          'MsgBox("hier=" & hier & " page=" & page & " offs=" & offs & " pageo=" & pageo)
          ' (1) 灰色メソッドでの目次スライドをまず追加
          ActivePresentation.Slides.Add pageo, ppLayoutText ' その前にスライド追加
          ActivePresentation.Slides(pageo).Shapes(1).TextFrame.TextRange=title ' title を書く
          ' 目次を書くプレースホルダー
          Set oTxtRng = ActivePresentation.Slides(pageo).Shapes(2).TextFrame.TextRange
          Dim toc As Variant
          ' +2 同じページに複数の項目の説明があることを想定する場合(同じ hierarchy でページ数が同じ)
          Set hsh=New Dictionary  ' +2
          Dim k As Variant                              ' +2
          'MsgBox("hier=" & hier & " page=" & page & " idx=" & iNcIdx(hier,page))
          For Each k In Split(iNcIdx(hier,page),vbNewLine)                ' +2
            If Not k = "" Then     ' 何故こんな値が入るのかよくわからない  '+2
              hsh.Add CInt(k),1                                            '+2
            End If                                                         '+2
          Next k                                                           '+2
          Dim idx As Integer
          idx = 0
          Dim sname As String
          sname  = C_GRAYTOC & "_" & title & "_" & CStr(hier) & "_" & CStr(page)
          Dim sname2 As String
          sname2=""
          For Each toc In Split(iNcToc(hier,page),vbNewLine) '一行ずつ分割
            With oTxtRng.Paragraphs(idx) '箇条書き1つずつ
              .Text = toc & vbNewLine
              ' すぐ次のページで説明する項目 = 黒で
              'If iNcIdx(hier,page) = idx Then ' +1 想定しない場合 '+1
              If hsh(idx) = 1 Then             ' 想定する場合   '+2
                .Font.Color.RGB=Blk
                sname2 = sname2 & "_" & toc
              Else
                .Font.Color.RGB=Gry   ' 関係ない項目 = 灰色で
              End If
            End With
            idx = idx + 1
          Next toc
          hsh.RemoveAll ' 内容を削除  '+2
          ActivePresentation.Slides(pageo).Name = sname & sname2
          ' (2) id が 0 であれば, その前に普通の目次を追加
          'If InStr(iNcIdx(hier,page),"0") Then  ' これだと id=10 の時も該当してしまう!! 2014/05/10
          If Left(iNcIdx(hier,page),1) = "0" Then
            'MsgBox("hier=" & hier & vbNewLine & "page=" & page & vbNewLine & "idx=" & iNcIdx(hier,page))
            ActivePresentation.Slides.Add pageo,ppLayoutText
            ActivePresentation.Slides(pageo).Shapes(1).TextFrame.TextRange=title
            ActivePresentation.Slides(pageo).Shapes(2).TextFrame.TextRange=iNcToc(hier,page)
            ActivePresentation.Slides(pageo).Name = sname
          End If
          offs = offs + iNcOff(hier,page)
        End If
      End If
    Next hier
  Next page
End Sub

' offset 値を計算します.
' offset(page) が無い場合, page = page + offset(UBound(offset)) を使います
Private Function oFFst (page As Integer, offset() As Integer) As Integer
  Dim n As Integer
  n = UBound(offset)

  Dim off As Integer
  off = 0
  If page > n Then
    off = offset(n)
  ElseIf Not IsNull(offset(page)) Then
    off = offset(page)
  End If
  oFFst = off
End Function

' offset 値及び exist 値(optional. 考慮するページ数を 0 or 1 で区別. exist(10)=1: 10 ページ目は考慮に入れること)
' を用いて, 階層目次(for 灰色目次)を新しくして書き出します.
Private Sub bufWrit_TOC(prnt As String, sBuf As Variant, delm As String, offset() As Integer, Optional pExist As Variant)

  If Sgn(offset) = 0 Then
    If Not IsNull(sBuf) And Not sBuf = "" Then
      prnt = prnt & sBuf & vbNewLine
    End If
    Exit Sub
  End If

  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine)

  Dim m As Integer
  m = UBound(aBuf)

  Dim j  As Integer
  Dim fst As Integer  ' 最初の書き出しのフラグ(書きだした行数)
  fst = 0
  For j = 0 To m
    If Left$(aBuf(j),1) = "*" Then
      Dim bf() As String
      bf = Split(aBuf(j),delm)    ' **(\t)目次(\t)(\t)ページ数(数値)
      Dim page As Integer
      page = CInt(bf(UBound(bf))) ' 最後のカラムがページ数(数値)
      'If pExist Is Nothing Then   ' pExist が定義されてないとき
      'If IsMissing(pExist) Then   ' これを通過して Else 以下でエラーとなることがある
      'If IsMissing(pExist) Or pExist Is Nothing Then ' これで大丈夫か? 2014/03/27
      ' pExist が配列として存在しているとき,
      ' If pExist Is Nothing Then ' これはエラーとなる Why?
      If IsMissing(pExist) Or IsEmpty(pExist) Or IsNull(pExist) Then ' 2014/04/06 ad hoc
        page = page + oFFst(page,offset()) ' page 数をずらす
        bf(UBound(bf)) = CStr(page)
        prnt = prnt & join(bf,delm) & vbNewLine
        fst= fst+1
      Else
        'If Not page > UBound(pExist) AndAlso pExist(page) = 1 Then    ' 2014/04/21 And -> AndAlso
        ' 2014/04/21 短絡評価ができない?
        If Not page > UBound(pExist) Then
          If pExist(page) = 1 Then
            page = page + oFFst(page,offset()) ' page 数をずらす
            ' ページの途中から始まっているときに, 階層がうまく記述できない
            ' 最初の書きだしの際に階層をたどって書いておく.
            ' 2014/04/25
            If fst = 0 Then
              Dim star As Integer
              star =Len(bf(0))
              If star > 1 Then      ' 階層が '*' でないとき
                Dim spx() As String ' 階層毎に文字列を入れておく箱をかくほ.
                ReDim spx(star)
                Dim jx As Integer
                For jx = 0 To j-1
                  Dim bfx() As String
                  bfx = Split(aBuf(jx),delm)
                  Dim starx As Integer
                  starx=Len(bfx(0))
                  If starx < star Then
                    bfx(UBound(bfx))=CStr(page)
                    spx(starx)=join(bfx,delm)
                  End If
                Next jx
                Dim s As Integer
                For s = 1 To star-1 ' 階層をたどって
                  If Not IsNull(spx(s)) Then
                    If Not spx(s) = "" Then
                      prnt = prnt & spx(s) & vbNewLine ' 最初の階層目次を書く
                    End If
                  End If
                Next s
              End If
            End If
            ' end 2014/04/25
            bf(UBound(bf)) = CStr(page)            
            prnt = prnt & join(bf,delm) & vbNewLine
            fst  = fst+1
          End If
        End If
      End If
    Else
      prnt = prnt & aBuf(j) & vbNewLine
    End If
  Next j

End Sub

' 階層目次の構造を読み込みます.
Private Sub bufRead_TOC(sBuf As Variant, iNcIdx() As String, iNcTtl() As String, iNcToc() As String, iNcOff() As Integer,delm As String)
  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine) ' 一行ずつ分割

  Dim m As Integer
  m = UBound(aBuf) ' m = 行数 -1 (0-based)

  ' 階層の数を数える(* の数を数える)
  Dim tBuf() As String
  Dim hier   As Integer
  Dim h,j    As Integer
  Dim p      As Variant

  hier = 1
  For j = 0 To m
    tBuf = Split(aBuf(j),vbTab)
    If left$(tBuf(0),1) = "*" Then  ' 2014/06/21 debug
      If hier< Len(tBuf(0)) Then hier=Len(tBuf(0))
    End If
  Next j

  ReDim iNcIdx(1 To hier,1 To 1)
  ReDim iNcTtl(1 To hier,1 To 1)
  ReDim iNcToc(1 To hier,1 To 1)
  ReDim iNcOff(1 To hier,1 To 1)

  Dim pmax As Integer
  pmax=0
  For h = 1 To hier ' 階層の数. 階層を一つずつ見ていく
    Dim title   As String
    Dim bstar   As Integer
    Dim bread   As Integer
    Dim tocs()  As String
    Dim pages() As Integer
    Dim ut      As Integer
    Dim up      As Integer
    bstar=0
    bread=0
    ut   =0
    up   =0
    title = C_TITLE  ' 2014/03/27 default title
    For j = 0 To m
      Dim page As Integer
      Dim star As Integer
      Dim toc  As String
      Dim bf() As String
      'Dim cf() As String
      If Left$(aBuf(j),1) = "*" Then
        bf   =Split(aBuf(j),delm)   ' tab split by delm (vbTab)
        page =CInt(bf(UBound(bf)))  ' 最後の要素 = ページ数
        star =Len(bf(0))            ' 2014/03/11 tab 区切りに変更
        toc  =bf(1)
        'cf   =Split(bf(0))          ' 左側を space split
        'star =Len(cf(0))
        'toc  =Mid(bf(0),star+2)
        If h = star Then
          If ut = 0 Then
            ReDim tocs(ut)
          Else 
            ReDim Preserve tocs(ut)
          End If
          tocs(ut)=toc
          ut = ut + 1
          If up = 0 Then
            ReDim pages(up)
          Else 
            ReDim Preserve pages(up)
          End If
          pages(up)=page
          up = up + 1
        End If
        If star < h And bread = 1 And Not pages(0) = 0 Then
          'MsgBox("make new toc: NOW LINE IS " & aBuf(j))
          Dim id As Integer
          id = 0
          'Dim xprnt As String
          'xprnt=""
          For Each p In pages
            If pmax < p Then
              pmax = p
              ReDim Preserve iNcIdx(1 To hier,1 To pmax)
              ReDim Preserve iNcTtl(1 To hier,1 To pmax)
              ReDim Preserve iNcToc(1 To hier,1 To pmax)
              ReDim Preserve iNcOff(1 To hier,1 To pmax)
            End If
            ' +2 同じページに複数の項目の説明があることを想定する場合
            If IsNull(iNcIdx(h,p)) Or iNcIdx(h,p) ="" Then  '+2
              iNcOff(h,p)= iNcOff(h,p) + 1                  '+2
              If id = 0 Then iNcOff(h,p)= iNcOff(h,p) + 1   '+2
            End If                                          '+2
            iNcIdx(h,p)=iNcIdx(h,p) & id & vbNewLine        '+2

            ' +1 想定しない場合
            ' 同じページに同じ hierarchy の説明を複数しないのであればこっちを使う方が良い
            ' +2, +1 入れ替え
            'iNcOff(h,p)= iNcOff(h,p) + 1                   '+1
            'If id = 0 Then iNcOff(h,p)= iNcOff(h,p) + 1    '+1
            'iNcIdx(h,p)= id                                '+1

            iNcTtl(h,p)=title
            iNcToc(h,p)=join(tocs,vbNewLine) ' 二次元配列面倒なので改行で

            'For Debug
            'xprnt = "title= " & title & vbNewLine
            'xprnt = xprnt & "id  = " & id & vbNewLine
            'xprnt = xprnt & "page= " & p  & vbNewLine
            'xprnt = xprnt & join(tocs,vbNewLine)
            'MsgBox(xprnt)
            ' End For Debug
            id = id + 1
          Next p
          ut=0
          up=0
          ReDim tocs(ut)
          ReDim pages(up)
        End If
        If h = star + 1 Then title = toc
        If h = star     Then bread = 1
        bstar = star
      Else
        bstar = 0
        If Not aBuf(j) = "" Then  ' 2014/03/27
          title = aBuf(j)
        End If
      End If
    Next j
    If Not IsNull(pages) And Not pages(0) = 0 Then
      id =0
      'xprnt="" ' For Debug
      For Each p In pages
        If pmax < p Then
          pmax = p
          ReDim Preserve iNcIdx(1 To hier,1 To pmax)
          ReDim Preserve iNcTtl(1 To hier,1 To pmax)
          ReDim Preserve iNcToc(1 To hier,1 To pmax)
          ReDim Preserve iNcOff(1 To hier,1 To pmax)
        End If
        iNcIdx(h,p)=id
        iNcTtl(h,p)=title
        iNcToc(h,p)=join(tocs,vbNewLine)
        iNcOff(h,p)= iNcOff(h,p) + 1
        If id = 0 Then iNcOff(h,p)= iNcOff(h,p) + 1

        'For Debug
        'xprnt = "title= " & title & vbNewLine
        'xprnt = xprnt & "id  = " & id & vbNewLine
        'xprnt = xprnt & "page= " & p  & vbNewLine
        'xprnt = xprnt & join(tocs,vbNewLine)
        'MsgBox(xprnt)
        ' End For Debug
        id = id + 1
      Next p
    End If
  Next h

  ' for Debug
  'Dim prnt As String
  'Dim g    As Integer
  'Dim q    As Integer
  'g = 21
  'q = 2
  'prnt = "DEBUG: page=" & g & " id = " & CStr(iNcIdx(q,g)) & vbTab & "title= " & iNcTtl(q,g) & vbNewLine
  'prnt = prnt & iNcToc(q,g)
  'MsgBox(prnt)
  ' End for Debug

End Sub

' text file を読み込んで hash に格納します
' String として読み込むだけ.
' nock が 1 の時は, ファイルの存在を確認しない(無ければそのまま終わる)
Private Sub fileRead2(sNotesFilePath As String, hBufs, Optional nock As Integer)

  If nock = 1 And Dir(sNotesFilePath) = "" Then
    Exit Sub
  End If

  ' ファイルを読み出す
  Dim iNotesFileNum As Integer
  iNotesFileNum = FreeFile()
  Open sNotesFilePath For Input As iNotesFileNum

  Dim MARK As String
  Dim sBuf As String
  Do Until EOF(iNotesFileNum)
    Line Input #iNotesFileNum,sBuf  ' sBuf の中に一行ずつ入れていく
    If sBuf = "" Then GoTo SKIP                 ' 空行は読み飛ばす
    If left$(sBuf, 1) = COMT_MARK Then GoTo SKIP  ' コメント行が来たら飛ばす
    If Left$(sBuf,2) = INFO_MARK And Right$(sBuf,2) = INFO_MARK And Len(sBuf)= Len(INFO_MARK)*2+3 Then '情報目印
      MARK = Mid$(sBuf,3,3)
    Else
      If Not IsNull(MARK) And Not MARK = "" Then
        If hBufs.Exists(MARK) Then
          hBufs(MARK) = hBufs(MARK) & vbNewLine & sBuf
        Else
          hBufs.Add MARK,sBuf
        End If
      End If
    End If
SKIP:
  Loop
  Close iNotesFileNum

End Sub

' text file 名を取得します.
' default では, hoge.pptx => hoge.txt
' ck = 1   の場合, ファイルが無ければ即プログラムを終了.
' mflg = 1 の場合, MsgBox で default の text file を使うかどうかを確認.
Private Sub notesFilePath (sNotesFilePath As String, sCurrentFolder As String, suffix As String, ck As Integer, mflg As Integer, Optional Fname As String)

  Dim delm As String
  delm = "\"   ' default (windows)

  Dim Op As Variant
  Op = Application.OperatingSystem
  If Op Like "Macintosh*" Then delm = ":"

  Dim sNotesFileName As String   ' text file 名
  ' 設定 text file
  sCurrentFolder = ActivePresentation.Path & delm
  ' mflg = 1 確認 MsgBox を表示する
  sNotesFileName = newFileName(suffix,"do you want to use the file",mflg,Fname)

  sNotesFileName = Replace(sNotesFileName,"\",delm)
  
  'MsgBox("Fname=" & Fname)
  'MsgBox("sNotesFileName=" & sNotesFileName)
  'MsgBox("sCurrentFolder=" & sCurrentFolder)
  
  If InStr(sNotesFileName, delm) = 0 Then
    sNotesFilePath = sCurrentFolder & sNotesFileName
  'ElseIf Left$(sNotesFileName,2) = "." & delm Then  ' 2014/04/11 削除
  '  sNotesFilePath = sCurrentFolder & Mid(Fname,2)
  ElseIf Left(sNoteFileName,1) = "." Then  ' 2014/04/11 やっぱりよくわからない.
    sNotesFilePath = sCurrentFolder & sNoteFileName
  Else
    sNotesFilePath = sNotesFileName  ' 絶対パスで書いてある場合
  End If
  ' is it there? quit if not
  'sNotesFilePath = sCurrentFolder & sNotesFileName
  'MsgBox("sNotesFilePath=" & sNotesFilePath)

  If ck = 1 Then  ' ファイルの存在を確認する場合 ck = 1
    If Len(Dir$(sNotesFilePath)) = 0 Then
      MsgBox (sNotesFilePath & " is missing")
      End
    End If
  End If

End Sub

' text file に 使う powerpoint ファイルとそのページ数を指定しておくと,
' これを動かすことで指定したスライドが今の powerpoint ファイルに取り込まれます.
' NEW: text file があれば, 取り込んだことによるページのずれ(offset)を考慮した
' 新しい統合された text file を作成します.
Public Sub pv_0000_collectPpt()

  ' 設定ファイルを読み込み,
  ' そこに書かれた他の ppt ファイルと指定したページをコピーする
  ' 最初の番号 (0 or 1) は, 0 のとき = スタイルはコピーしない. 1 のとき = スタイルもコピー)
  ' ::+++::
  ' 0   hoge.ppt    1-3,10,45,44
  ' 1   fuga.pptx   20,30,50-60
  ' 1   hoge.ppt    45

  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  Set hBufs = New Dictionary
  Call fileRead2(sNotesFilePath,hBufs)

  ' ファイル情報の取得
  ' Desgns(N)  コピーの方法(0: スタイルは target ppt 1: スタイルを保持してコピー)
  ' Fnames(N)  用いる ppt ファイルのリスト
  ' Pdummy(N)  ページ数を書いたリスト
  Dim   Desgns() As Integer
  Dim   Fnames() As String
  Dim   Pagesc() As String
  'vbTab で固定("," に変更はできない. Pagesc が"," 区切りなので)
  Call bufRead_File(hBufs(CPPT_MARK),Desgns(),Fnames(),Pagesc(),C_TBB)

  ' 今開いているスライド
  Dim pTo  As Presentation
  Set PTo = Application.ActivePresentation
  
  ' 保存する資料のファイル名
  Dim sPptFileName As String
  spptFileName = newFileName("_intg.pptm","do you want to save the file as ",1)

  ' 別名で上書き保存
  ' ここからはこのファイルが Active となるので注意する.
  If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

  Dim m As Integer
  m = UBound(Fnames)

  Dim sttIns() As Integer
  ReDim sttIns(m)

  ' ppt ファイルの取り込み
  Dim k As Integer
  For k = 0 To m
    Dim pFr As Presentation ' ppt ファイルを開きます.
    Set pFr = Presentations.Open(FileName:=Fnames(k),ReadOnly:=msoFalse)
    Dim sttIn As Integer
    'sttIn = ActivePresentation.Slides.Count + 1 ' paste されるページの最初のページ
    'sttIn    = pTo.Slides.Count + 1  ' paste されるページの最初のページ
    sttIn    = pTo.Slides.Count  ' paste されるページの最初のページ  ' 2014/03/24
    'MsgBox("sttIn = " & sttIn)
    sttIns(k)=sttIn
    If Desgns(k) = 0 Then
      Call copySlide(pFr,pTo,Pagesc(k),C_CMM)
    Else
      Call copySlide_Fmt(pFr,pTo,Pagesc(k),C_CMM)
    End If
  Next k

  ' 統合 text file の作成
  Dim sItgFilePath   As String
  'Call notesFilePath(sItgFilePath,sCurrentFolder,"_intg.txt",0,1)
  Call notesFilePath(sItgFilePath,sCurrentFolder,".txt",0,1) ' 2014/04/06
  Call offsetInteg(Fnames(),Pagesc(),C_CMM,sttIns(),sItgFilePath)

  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

End Sub

' いくつかの text file について offset 補正し,
' 統合テキストファイルを作成します.
Private Sub offsetInteg (Fname() As String, Pagesc() As String, delm As String, sttIns() As Integer, sFilePath As String)
  Set nBufs = New Dictionary ' hash
  Dim m As Integer
  m = UBound(Fname)
  For k = 0 To m
    Call offset_Intg(nBufs,Fname(k),Pagesc(k),delm,sttIns(k))
  Next k
  Call hashPrint(nBufs,sFilePath)
End Sub

' 取り込むページのリストと現在のページ数から,
' offset() の値と pExit() を計算します.
Private Sub offset_Intg(nBufs,Fname As String, Pagesc As String, delm As String, sttIn As Integer)

  Dim Pages() As Integer
  Call sList2iList(Pagesc,Pages(),delm)

  Dim npage As Integer
  Dim p     As Variant
  For Each p In Pages
    If npage < p Then npage = p
  Next p

  ' page offset
  ' new = oFFst(page,offset())
  Dim offset() As Integer
  ReDim offset(npage)
  Dim pg As Integer
  pg = sttIn + 1
  For Each p In Pages
    offset(p) = pg - p
    pg = pg + 1
  Next p

  ' pExist(関係あるページ数) = 1
  ' pExist(ないページ数)     = 0
  Dim pExist() As Integer
  ReDim pExist(npage)
  Dim i As Integer
  For i = 0 To npage
    pExist(i)=0
  Next i
  For Each p In Pages
    pExist(p)=1
  Next p

  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,0,Fname)
  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs,1)  ' 最後の引数 =1 ファイル無くてもそのまま進む.
  Dim Prng As String
  Call joinInt(Prng,pExist(),",")
  Call offsetHash(hBufs,nBufs,offset(),pExist)

End Sub

' collectPpt 用の情報を書き出します.
Private Sub bufWrit_File(prnt As String, sBuf As Variant, delm As String, offset() As Integer, Optional pExist As Variant)

  If Sgn(offset) = 0 Then
    If Not IsNull(sBuf) And Not sBuf = "" Then
      prnt = prnt & sBuf & vbNewLine
    End If
    Exit Sub
  End If

  Dim   Desgns() As Integer
  Dim   Fnames() As String
  Dim   Pagesc() As String
  Call bufRead_File(sBuf,Desgns(),Fnames(),Pagesc(),delm)

  Dim m As Integer
  m = UBound(Fnames)

  Dim k  As Integer
  For k = 0 To m
    Dim sForm As String
    sForm=""  ' Why ?
    Call sList2OFFSETsForm(Pagesc(k),sForm,C_CMM,offset(),pExist)
    If sForm <> "" Then  ' 2014/04/29 ページ数が無い場合の置換文字列は書かない
      prnt = prnt & Desgns(k) & delm & Fnames(k) & delm & sForm & vbNewLine
    End If
  Next k

End Sub

' collectPpt 用の情報を読み込みます
Private Sub bufRead_File(sBuf As Variant, Desgns() As Integer, Fnames() As String, Pagesc() As String, delm As String)
  Dim Op As Variant
  Op = Application.OperatingSystem

  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine) ' 一行ずつ分割

  Dim fnum     As Integer
  Dim Pdummy() As String
  fnum = 0
  Dim tBuf As Variant
  For Each tBuf In aBuf
    ReDim Preserve Desgns(fnum)
    ReDim Preserve Fnames(fnum)
    ReDim Preserve Pdummy(fnum)
    Dim buf() As String
    buf = Split(tBuf,delm) ' split by vbTab(固定)
    ' 1. スタイルをコピーするかどうか (0 or 1)
    Desgns(fnum)=CInt(buf(0))
    ' 2. ppt file name
    Dim Fname As String
    ' 2014/06/27 関数にした
    Fname=filePath(buf(1))
'    Fname = buf(1)
'    If Op Like "Macintosh*" Then
'      Fname = Replace(Fname,"\",":")
'      If InStr(Fname, ":") = 0 Then
'        Fname = ActivePresentation.Path & ":" & Fname
'      ElseIf Left$(Fname,2) = ".:" Then
'        Fname = ActivePresentation.Path & ":" & Mid(Fname,2)
'      End If
'    Else
'      ' もしFname がファイル名だけの場合は, カレントディレクトリにあるとみなして
'      ' 絶対パスにする.
'       'あるいは "." から始まる場合には相対パスで書いてあると考えて
'      ' 今いるディレクトリ名を追記して絶対パスにする
'      ' それ以外の場合は, 絶対パスで書いてあるとかんがえる.
'      If InStr(Fname, "\") = 0 Then
'        Fname = ActivePresentation.Path & "\" & Fname
'      'ElseIf Left$(Fname,2) = ".\" Then   ' 2014/04/11 削除
'      '  Fname = ActivePresentation.Path & "\" & Mid(Fname,2)
'      ElseIf Left(Fname,1) = "." Then ' 2014/04/11 よくわからない.
'        Fname = ActivePresentation.Path & "\" & Fname
'      End If
'    End If
    Fnames(fnum) = Fname
    ' 3. page number
    Pdummy(fnum)=buf(UBound(buf))
    fnum = fnum + 1
  Next tBuf

  Dim k  As Integer
  Dim fn As Integer
  fn = UBound(Fnames)
  ReDim Pagesc(fn)
  For k = 0 To fn
    Dim Pptf As Presentation
    Set Pptf = Presentations.Open(FileName:=Fnames(k),ReadOnly:=msoFalse) ' 使うスライドをとりあえず開く
    Dim defList() As Integer
    Call mkDefList(defList(),Pptf)
    Dim npage As Integer
    npage = Pptf.Slides.Count
    Pptf.Close
    ' ad hoc 2014/04/11
    'Pdummy(k) = "2-" のように, "-" で終わっている場合には, npage を追加
    'Pdummy(k) = "2-18"
    If Right(Pdummy(k),1) = "-" Then
      Pdummy(k) = Pdummy(k) & CStr(npage)
    End If
    Call sForm2sList(Pdummy(k),Pagesc(k),deflist(),C_CMM)
  Next k
End Sub

' slide をコピーします.
Private Sub copySlide(pFr As Presentation, pTo As Presentation, Pagesc As String, delm As String)
  Dim Pages() As Integer
  Call sList2iList(Pagesc,Pages(),C_CMM)
  ' copy slide (as usual)
  pFr.Slides.Range(Pages).Copy
  pTo.Slides.Paste
  pFr.Close
End Sub

' slide のスタイルを元の slide の状態に保ったままコピーします
Private Sub copySlide_Fmt(pFr As Presentation, pTo As Presentation, Pagesc As String, delm As String)

  Dim Pages() As Integer
  Call sList2iList(Pagesc,Pages(),C_CMM)

  ' keep formatting copy
  ' http://img2.tapuz.co.il/forums/1_88724584.txt
  Dim q  As Integer
  For q = 0 To UBound(Pages) ' 1 page ずつコピーする
    Dim sFr As Slide
    Set sFr = pFr.Slides(Pages(q))
    sFr.Copy
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


' 左下の countbox を作成します.
' 無視ページ "::---::"
' 及び gray 目次ページは
' 無視します.
Public Sub pv_0011_pageCountsBox2()

  ' text file を読み込んで page 番号を付けないものを指定
  ' page 番号を付けないもの: "::---::" で書かれたページ
  ' スライド名が grayTOC で始まるもの

  ' 色設定
  Dim Red, Blk, Gry, Wht, Col
  Red = RGB(255, 0, 0)
  Blk = RGB(0, 0, 0)
  Gry = RGB(70, 70, 70)
  Wht = RGB(255, 255, 255)
  Col = RGB(0, 150, 0)
  'Red = RED
  'Blk = BLACK
  'Gry = GRYA2
  'Wht = WHITE
  'Col = GREEN

  ' 場所設定
  Dim lef As Integer
  Dim top As Integer
  Dim wid As Integer
  Dim hei As Integer
  Dim stp As Integer
  Dim wei As Integer
  lef = -9
  top = 532
  wid = 8
  hei = 8
  stp = 8.5
  wei = 0.5

  Dim iSkip() As Integer
  Dim npage0  As Integer  ' スライド枚数
  Dim npage   As Integer  ' skip を除いたスライド枚数
  Call skipSlides(iSkip(),npage0,npage)
  
  Dim ShapeList() As String
  ReDim ShapeList(npage)
  
  Dim page As Integer
  page = 0
  Dim i As Integer
  Dim Shp
  For i = 1 To npage0
    If Not iSkip(i) = 1 Then
      page = page + 1
      'ActivePresentation.Slides(i).Select
      Dim j As Integer
      For j = 1 To npage
        Set Shp = ActivePresentation.Slides.Item(i).Shapes.AddShape(msoShapeRectangle, lef + (j * stp), top, wid, hei)
        'Shp.Select
        '2012/04/24 initialize fill pattern
        Set oSp=ActivePresentation.Slides(i).Shapes(ActivePresentation.Slides(i).Shapes.Count)
        oSp.Fill.Solid
        oSp.Line.Weight = wei
        ' 枠線
        If (j Mod 10 = 0) Then
          oSp.Line.ForeColor.RGB=Red
        Else
          oSp.Line.ForeColor.RGB=Blk
        End If
        ' 塗りつぶし
        If (j < page) Then
          oSp.Fill.ForeColor.RGB=Gry
        ElseIf (j = page) Then
          oSp.Fill.ForeColor.RGB=Col
        Else
          oSp.Fill.ForeColor.RGB=Wht
        End If
        ShapeList(j) = Shp.Name
      Next
      ' グループ化して名前をつける
      ActivePresentation.Slides(i).Shapes.Range(ShapeList()).Group.Name = C_COUNTBOX
    End If
  Next
  ActivePresentation.Slides(1).Select
  ' メモリ解放
End Sub

Public Sub removeCountsBox()
  Dim SlideObj   As Slide
  Dim ShapeObj   As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoGroup Then            ' グループ
        If ShapeObj.Name = C_COUNTBOX Then      ' 名前が "countbox"
          ShapeObj.Delete
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub


' ページ番号を付けます
' 無視ページ "::---::"
' 及び gray 目次ページは
' 無視します.
Public Sub pv_0012_pageNum2()

  ' text file を読み込んで page 番号を付けないものを指定
  ' page 番号を付けないもの: "::---::" で書かれたページ
  ' スライド名が grayTOC で始まるもの

  ' 場所指定
  Dim xPos As Integer
  Dim yPos As Integer
  Dim wid  As Integer
  Dim hei  As Integer
  'xPos = 686
  'yPos = 527
  'wid  = 40
  'hei  = 28.875
  xPos = C_XPOSB
  yPos = C_YPOSB
  wid  = C_WIDEB
  hei  = C_HIGHB

  ' font
  Dim Fname As String
  Dim Ename As String
  Dim Fsize As Integer
  Fname = C_FNAME
  Ename = C_ENAME
  Fsize = C_PSIZE

  Dim iSkip() As Integer
  Dim npage0  As Integer  ' スライド枚数
  Dim npage   As Integer  ' skip を除いたスライド枚数
  Call skipSlides(iSkip(), npage0, npage) ' iSkip(page)=1 のとき, page 番目のスライドをスキップする

  Dim page As Integer
  page = 0
  Dim i As Integer
  Dim Shp
  For i = 1 To npage0
    If Not iSkip(i) = 1 Then
      page = page + 1
      ActivePresentation.Slides(i).Select
      ' i 枚目のスライド
      Set Shp = ActivePresentation.Slides.Item(i).Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, wid, hei)
      Shp.Select
      ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=i, Length:=0).Select
      With ActiveWindow.Selection.TextRange
        .Text = page & "/" & npage
        With .Font
          .NameAscii = Fname
          .NameFarEast = Ename
          .NameOther = Fname
          .Size = Fsize
          .Bold = msoFalse
          .Italic = msoFalse
          .Underline = msoFalse
          .Shadow = msoFalse
          .Emboss = msoFalse
          .BaselineOffset = 0
          .AutoRotateNumbers = msoTrue
          .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0)
        End With
      End With
      Shp.Name = C_PAGENUM  ' textbox に名前をつけておく
    End If
  Next i
End Sub

Public Sub removePageNum()
  Dim SlideObj As Slide
  Dim ShapeObj As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides          ' 各スライド
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1   ' 各シェイプ
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoTextBox Then              ' TextBox である
        If ShapeObj.Name = C_PAGENUM Then             ' 名前が counttxt である.
          ShapeObj.Delete                             ' オブジェクトを消去する.
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub

' 無視するページを計算します
Private Sub skipSlides(iSkip() As Integer, ByRef npage0 As Integer, ByRef npage As Integer)

  ActivePresentation.Slides(1).Select     ' 一枚目のスライドを選択
  npage0 = ActivePresentation.Slides.Count ' スライドの総数(全部)

  ReDim iSkip(npage0)
  Dim isk As Integer
  isk = 0

  ' "::---::" フラグをテキストファイルからとってくるかどうか
  If MsgBox("do you want to use text file for skip ?",vbYesNo) = vbYes Then
    ' text file(path) の取得
    Dim sNotesFilePath As String
    Dim sCurrentFolder As String
    Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",0,1)
    Set hBufs = New Dictionary ' 読み込んだ文字列
    Call fileRead2(sNotesFilePath,hBufs)
    Dim nSids()  As Integer
    Call bufRead_nList(hBufs(SKIP_MARK),nSids(),C_CMM) ' vbTab -> ","
    Dim k As Integer
    If Not Sgn(nSids) = 0 Then  ' 2014/03/27
      For k = 0 To UBound(nSids)
        iSkip(nSids(k))=1
        isk = isk + 1
      Next k
    End If
  End If

  ' ページ数としてカウントしない目次の枚数を数える
  Dim i As Integer
  For i = 1 To npage0
    If InStr(ActivePresentation.Slides(i).Name, C_GRAYTOC) Then
      iSkip(i)=1
      isk = isk + 1
    End If
  Next i
  npage = npage0 - isk
End Sub

' 各々のスライド中に目次を作成し,
' 現在どこにあるかを明示します.
Public Sub pv_0013_pageListContents2()

  ' ppt の目次を自動で書き出す.
  ' 各々のページ(右上)に, 今やってる目次の位置を黒い色で,
  ' その他の目次の位置を灰色で書き出す
  ' 目次は, powerpoint ファイル名と同名のテキストファイルに書いておく.
  ' 例: hoge.ppt の目次を hoge.txt に書いておく.
  ' hoge.txt の中身:
  ' ::###::
  ' 1. 背景              1-3
  '    1.1. はじめに     2
  '    1.2. 歴史         3
  ' 2. 目的              4-6
  '...
  ' ::$$$::
  ' 1,2,3,4
  ' hoge.txt の中身終り
  '
  ' hoge.txt の format:
  '::###::
  ' (書き出す内容)\t\t\t(ページ数),(ページ数)-(ページ数)
  '::$$$::
  '(ページ数),(ページ数)-(ページ数)...     <- 目次を書き出さないページ数
  ' ...
  '
  '
  ' 2013/05/16 option (もし以下の情報があればこれを使う. 無ければ MsgBox で聞いてくる)
  ' 2014/03/01 format の変更
  ' ::!!!::
  '(目次を置く場所のx座標)\t(目次textboxの幅)\t(フォントサイズ)

  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  ' file reading
  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)

  ' 目次として書きだす内容とページ数
  Dim Conts()  As String
  Dim Pagesc() As String
  If Not IsNull(hBufs(TOCS_MARK)) And Not hBufs(TOCS_MARK) = "" Then
    Call bufRead_pList(hBufs(TOCS_MARK),Conts(),Pagesc(),C_TBB)
  ElseIf Not IsNull(hBufs(GTOC_MARK)) And Not hBufs(GTOC_MARK) = "" Then
    Call bufRead_TOC2pList(hBufs(GTOC_MARK),Conts(),Pagesc(),C_TBB)
  End If

 ' 目次を書き出さないスライド番号(1-based)
  Dim nSids()  As Integer
  Call bufRead_nList(hBufs(NOWT_MARK),nSids(),C_CMM) ' vbTab -> ","

  ' 書きだす場所
  Dim xPos0  As Integer
  Dim wid0   As Integer
  Dim Fsize0 As Integer
  Dim sizFlg As Integer
  sizFlg = 0
  If Not IsNull(hBufs(IFNT_MARK)) And Not hBufs(IFNT_MARK) = "" Then
    Call bufRead_siz(hBufs(IFNT_MARK),xPos0,wid0,Fsize0)
    sizFlg = 1
  End If

  Dim xPos  As Integer
  Dim yPos  As Integer
  Dim wid   As Integer
  Dim hei   As Integer
  Dim Fsize As Integer
  'xPos  = 590
  'xPos  = 610
  'yPos  = 5
  'wid   = 125
  'wid   = 140
  hei   = 20    ' 関係ない
  'Fsize = 10
  xPos  = C_XPOSR
  yPos  = C_YPOSR
  wid   = C_WIDER
  Fsize = C_FSIZE

  ' 書き出す場所, 幅, フォントサイズの変更
  Call siz_MsgBox(sizFlg,xPos,wid,Fsize,xPos0,wid0,Fsize0)

  ' スライドへ書き出し
  Call write_TOC(xPos,yPos,wid,hei,Fsize,nSids(),Conts(),Pagesc(),C_CONTENT,1)

End Sub

' 実際に目次を書き出します
Private Sub write_TOC(xPos As Integer,yPos As Integer,wid As Integer,hei As Integer,Fsize As Integer,nSids() As Integer,Conts() As String,Pagesc() As String,sName As String, aFlg As Integer)

  Dim iSkip() As Integer
  Dim npage0  As Integer  ' スライド枚数
  Dim npage   As Integer  ' skip を除いたスライド枚数
  Call skipSlides(iSkip(),npage0,npage)

  Dim fSld() As Integer ' 書きだすスライド番号 i fSld(i)=1
  Call checkw_slde(nSids(),fSld(),npage0)

  Dim cnum As Integer   ' 目次項目数
  cnum = UBound(Conts)

  Dim Blk() As Variant
  Dim Gry() As Variant
  Blk = Array(0, 0, 0)
  Gry = Array(180, 180, 180)

  Dim Fname, Ename
  Fname = C_FNAME
  Ename = C_SNAME

  Dim pCnt() As String
  ReDim pCnt(cnum)
  Dim nSl() As Integer
  ReDim nSl(cnum)
  Dim jc As Integer
  For jc = 0 To cnum
    Dim Pages() As Integer
    Call sList2iList(Pagesc(jc),Pages(),C_CMM)
    Dim iNc() As Integer
    ReDim iNc(npage0)
    Dim ic As Integer
    For ic = 0 To npage0
      iNc(ic)=0
    Next ic
    Dim g As Integer
    Dim k As Integer
    k = 0
    For g = 0 To UBound(Pages)
      If Not iSkip(Pages(g)) = 1 Then
        k = k + 1
        iNc(Pages(g))= k
      End If
    Next g
    nSl(jc) = k
    Dim sList As String
    Call iList2sList(iNc(),sList,C_CMM)
    pCnt(jc)=sList
  Next jc

  Dim Shp
  Dim i As Integer
  For i = 1 To npage0
    If fSld(i) = 1 And Not iSkip(i) = 1 Then
      ActivePresentation.Slides(i).Select
      ' i 枚目のスライド
      Set Shp = ActivePresentation.Slides.Item(i).Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, wid, hei)
      Shp.Select
      ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=i, Length:=0).Select
      Dim j As Integer
      For j = 0 To cnum
        Dim bWpg() As Integer
        Call sList2iList(pCnt(j),bWpg(),C_CMM)
        Dim ncnt As String
        ncnt = ""
        If Not bWpg(i) = 0 Then ncnt = " (" & bWpg(i) & "/" & nSl(j) & ")"
        Dim cont As String
        If j = cnum Then
          cont = Conts(j) & ncnt
        Else
          cont = Conts(j) & ncnt & vbNewLine ' 最後以外は改行を入れる
        End If

        Dim Col() As Variant
        Col = Gry
        If Not bWpg(i) = 0 Then
          Col = Blk
        End If
        Dim wFlg As Integer
        wFlg = 0
        If aFlg = 1 Then
          wFlg = 1
        Else
          If Col(0) = Blk(0) And Col(1) = Blk(1) And Col(2) = Blk(2) Then wFlg = 1
        End If
        If wFlg = 1 Then
          With ActiveWindow.Selection.TextRange
            .Text = cont
            With .Font
              .NameAscii = Fname
              .NameFarEast = Ename
              .NameOther = Fname
              .Size = Fsize
              .Bold = msoFalse
              .Italic = msoFalse
              .Underline = msoFalse
              .Shadow = msoFalse
              .Emboss = msoFalse
              .BaselineOffset = 0
              .AutoRotateNumbers = msoTrue
              .Color.RGB = RGB(Red:=Col(0), Green:=Col(1), Blue:=Col(2))
            End With
          End With
        End If
      Next
      Shp.ZOrder (msoSendToBack)  ' 最下層に置く
      Shp.Name = sName
    End If
  Next

End Sub

Private Sub bufRead_TOC2pList(sBuf As Variant,Conts() As String, Pagesc() As String, delm As String)
  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine) ' 一行ずつ分割

  Dim m As Integer
  'm = UBound(aBuf)-1 ' m = 行数 -2 (0-based) ' 最初の行(* がない)は考えない.
  m = UBound(aBuf)  ' 最初の行が無いときがあるので 2013/03/29

  Dim hier   As Integer
  Dim tBuf() As String
  Dim j      As Integer
  hier = 1
  For j = 0 To m
    tBuf = Split(aBuf(j))
    If hier< Len(tBuf(0)) Then hier=Len(tBuf(0))
  Next j

  Set hbef = New Dictionary
  Dim stt() As Integer
  Dim edd() As Integer
  Dim sForm() As String
  For j = 0 To m
    Dim buff As String
    'buff = aBuf(j+1)
    buff = aBuf(j)    ' 2013/03/29
    If Left$(buff,1) = "*" Then
      ' 動的に変更 2013/03/29
      ReDim Preserve sForm(j) 
      ReDim Preserve stt(j)
      ReDim Preserve edd(j)
      ReDim Preserve Conts(j)
      ReDim Preserve Pagesc(j)
      Dim page As Integer
      Dim star As Integer
      Dim toc  As String
      Dim bf() As String
      'Dim cf() As String
      bf        =Split(buff,delm)      ' tab split by delm (vbTab)
      page      =CInt(bf(UBound(bf)))  ' 最後の要素 = ページ数
      star      =Len(bf(0))            ' 2014/03/11 tab 区切りに変更
      toc       =bf(1)
      'cf        =Split(bf(0))          ' 左側を space split
      'star      =Len(cf(0))            ' 階層の数
      'toc       =Mid(bf(0),star+2)     ' 目次

      Conts(j)  =toc
      stt(j)    =page
      Dim st As Integer
      For st = star To hier
        If hbef.Exists(st) Then
          If Not stt(hbef(st)) = page Then
            edd(hbef(st))=page-1
          Else
            edd(hbef(st))=C_EXC
          End If
          hbef(st) = j
        Else
          hbef.Add st,j
        End If
      Next st
    End If
  Next j

  ' last
  page = ActivePresentation.Slides.Count + 1
  star = 1
  For st = star To hier
    If hbef.Exists(st) Then
      If Not stt(hbef(st)) = page Then
        edd(hbef(st))=page-1
      Else
        edd(hbef(st))=C_EXC
      End If
      hbef(st) = j
    Else
      hbef.Add st,j
    End If
  Next st

  For j = 0 To m
    'If Not edd(j) = C_EXC Then
    If Not edd(j) = C_EXC And stt(j) < edd(j) Then   ' ad hoc 2014/03/27
      sForm(j) = CStr(stt(j)) & "-" & CStr(edd(j))
    Else
      sForm(j) = CStr(stt(j))
    End If
  Next j

  Dim defList() As Integer
  Call mkDefList(defList())
  For j = 0 To m
    Call sForm2sList(sForm(j),Pagesc(j),defList(),C_CMM)
  Next j

End Sub

' 書きだすスライド番号 k: fSld(k) = 1
' nSids(i)=1: i 番目は書きださない
Private Sub checkw_slde(nSids() As Integer, fSld() As Integer, npage As Integer)
  ' 書き出さないスライド nSids() = 例: (1,2,5)
  ' i 枚目のスライドに目次を書き出す: fSld(i)=1
  ' 書き出さない: 例: fSld(1,2,5)=0
  ReDim Preserve fSld(1 To npage)
  Dim i As Integer
  For i = 1 To npage
    fSld(i) = 1       ' 基本は 1
  Next i
  If Not Sgn(nSids) = 0 Then
    Dim j As Variant
    For Each j In nSids
      fSld(j) = 0       ' flag が立っている = 0
    Next j
  End If
End Sub

' 書きだすスライド番号 k: fSld(k) = 1
' nSids(i)=1: i 番目を書きだす
Private Sub checkd_slde(nSids() As Integer, fSld() As Integer, npage As Integer)
  ReDim Preserve fSld(1 To npage)
  Dim i As Integer
  For i = 1 To npage
    fSld(i) = 0      ' 基本は 0
  Next i
  If Not Sgn(nSids) = 0 Then
    Dim j As Variant
    For Each j In nSids
      fSld(j) = 1      ' flag が立っている = 1
    Next j
  End If
End Sub

' (目次内容)(\t)(ページ数のリスト)
' を書きだす
' (ページ数のリスト)には, offset()及び, pExist() で今考えているページのみで計算しなおしたもの
Private Sub bufWrit_pList(prnt As String, sBuf As Variant, delm As String, offset() As Integer, Optional pExist As Variant)

  If Sgn(offset) = 0 Then
    If Not IsNull(sBuf) And Not sBuf = "" Then
      prnt = prnt & sBuf & vbNewLine
    End If
    Exit Sub
  End If

  Dim Conts()  As String
  Dim Pagesc() As String
  Call bufRead_pList(sBuf,Conts(),Pagesc(),delm)

  Dim m As Integer
  m = UBound(Conts)

  Dim k As Integer
  For k = 0 To m
    Dim sForm As String
    sForm = "" ' Why ?
    Call sList2OFFSETsForm(Pagesc(k),sForm,C_CMM,offset(),pExist)
    If sForm <> "" Then '2014/03/24 ページ数が無い目次は書かないことにする.
      prnt = prnt & Conts(k) & delm & sForm & vbNewLine
    End If
  Next k

End Sub

'(目次内容)(\t)(ページ数のリスト) を読み込む
' (ページ数のリスト)は sList 形式 = "1,2,4,5,6,7,8" のような感じで持っておく.
Private Sub bufRead_pList(sBuf As Variant, Conts() As String, Pagesc() As String, delm As String)

  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine)

  ' * は使わない(dummy)
  Dim defList(0) As Integer
  defList(0)=C_EXC

  Dim m As Integer
  m = UBound(aBuf)
  ReDim Conts(m)
  ReDim Pagesc(m)
  Dim k As Integer
  For k = 0 To m
    Dim buf() As String
    buf = Split(aBuf(k),delm)
    Conts(k) =buf(0)
    Call sForm2sList(buf(UBound(buf)),Pagesc(k),defList(),C_CMM)
  Next k

End Sub

' (ページ数のリスト)を書きだす
Private Sub bufWrit_nList(prnt As String, sBuf As Variant, delm As String, offset() As Integer, Optional pExist As Variant)

  If Sgn(offset) = 0 Then
    If Not IsNull(sBuf) And Not sBuf = "" Then
      prnt = prnt & sBuf & vbNewLine
    End If
    Exit Sub
  End If

  Dim nSids()  As Integer
  Call bufRead_nList(sBuf,nSids(),delm)

  Dim sForm As String
  Call iList2OFFSETsForm(nSids(),sForm,delm,offset(),pExist)
  prnt = prnt & sForm & vbNewLine

End Sub

' (ページ数のリスト)を読み込む
Private Sub bufRead_nList(sBuf As Variant, nSids() As Integer,delm As String)
  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine)

  Dim sForm As String
  sForm = join(aBuf,delm)

  ' * は使わない
  Dim defList(0) As Integer
  defList(0)=C_EXC

  Call sForm2iList(sForm,nSids(),delm,defList)

End Sub

' 書き込む位置, font 情報を書きだす
Private Sub bufWrit_siz(prnt As String, sBuf As Variant)
  prnt = prnt + sBuf
End Sub

Private Sub bufRead_siz(sBuf As Variant, xPos0 As Integer, wid0 As Integer, Fsize0 As Integer)
  ' 注: format の変更
  ' ::!!!::
  ' xPos0 \t wid0 \t Fsize0   <- sBuf

  Dim aBuf() As String
  aBuf = Split(sBuf,C_TBB)
  xPos0 =CInt(aBuf(0))
  wid0  =CInt(aBuf(1))
  Fsize0=CInt(aBuf(2))
End Sub

' font size 等の変更を MsgBox で
Private Sub siz_MsgBox(flg As Integer, x As Integer ,w As Integer ,f As Integer, x0 As Integer,w0 As Integer, f0 As Integer)
  If flg = 1 Then
    x=x0
    w=w0
    f=f0
  Else
    If MsgBox("x position= " & x & " textbox width =" & w & " font size = " & f & ": are they OK ?", vbYesNo) = vbNo Then
      x = InputBox("input x position of the textbox", "x の値が小さい=左側から書き出し", x)
      w = InputBox("input textbox width", "textbox width", w)
      f = InputBox("input font size", "font size", f)
    End If
  End If
End Sub

Public Sub removeListContents()
  Dim SlideObj As Slide
  Dim ShapeObj As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides          ' 各スライド
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1   ' 各シェイプ
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoTextBox Then           ' TextBox である
        If ShapeObj.Name = C_CONTENT Then            ' 名前が tcontents である.
          ShapeObj.Delete                          ' オブジェクトを消去する.
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub


' 各々のスライド中に目次を作成し,
' 現在どこにあるかを明示します.
Public Sub pv_0014_pageListHierarchy2()

  ' 2013/05/16
  ' ppt の階層構造を書きだす
  ' code from  pageListContents() ほとんどこぴぺで作成(何とかする)
  ' 2014/04/09 仕様の変更
  ' いままで
  ' pageListHierarchy(^^^) -> pageListContents(###) -> grayTOC(|||) の優先順位
  ' これを以下に変更
  '  pageListHierarchy(^^^) -> grayTOC(|||) -> pageListContents(###) の優先順位
  ' &&& を目印に書き出さないページを指定する.
  ' ??? を目印に, 書きだす場所(x座標), textbox の幅, font size を指定する(option)

  '
  ' 各々のページ(左上)に, 今やってる目次の位置(階層)を書きだす.
  ' 目次は, powerpoint ファイル名と同名のテキストファイルに書いておく.
  ' 例: hoge.ppt の目次を hoge.txt に書いておく.
  ' 何もなければ, pageListContents() の情報をそのまま使う.
  ' hoge.txt の中身:
  ' ::^^^::
  ' 1. 背景              1-3
  '    1.1. はじめに     2
  '    1.2. 歴史         3
  ' 2. 目的              4-6
  ' ::&&&::
  ' 1-4,8
  ' hoge.txt の中身終り
  '
  ' hoge.txt の format:
  '::^^^::
  ' (書き出す内容)\t\t\t(ページ数),(ページ数)-(ページ数)
  '::&&&::
  '(ページ数),(ページ数)-(ページ数)...     <- 目次を書き出さないページ数
  '
  '2013/05/16 option (もし以下の情報があればこれを使う. 無ければ MsgBox で聞いてくる)
  '2014/03/01 format 変更
  '::???::
  '(目次を置く場所のx座標)\t(目次textboxの幅)\t(フォントサイズ)
  '例:
  '::???::
  '10   130    8
  '
  ' ^^^ が無い場合には, gray 目次(|||)を用いる.
  ' hoge.txt の中身:
  ' ::|||::
  ' *     1. 背景            1-10
  ' **    1.1. はじめに      1-3
  ' ***   1.1.1. 背景(1)     1,2
  
  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  ' file reading
  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)

  ' 目次として書きだす内容とページ数
  Dim Conts()  As String
  Dim Pagesc() As String
  ' 2041/04/09 順序の変更
  If Not IsNull(hBufs(TOHI_MARK)) And Not hBufs(TOHI_MARK)="" Then
    Call bufRead_pList(hBufs(TOHI_MARK),Conts(),Pagesc(),C_TBB)
  ElseIf Not IsNull(hBufs(GTOC_MARK)) And Not hBufs(GTOC_MARK) = "" Then
    Call bufRead_TOC2pList(hBufs(GTOC_MARK),Conts(),Pagesc(),C_TBB)
  ElseIf Not IsNull(hBufs(TOCS_MARK)) And Not hBufs(TOCS_MARK) = "" Then 
    Call bufRead_pList(hBufs(TOCS_MARK),Conts(),Pagesc(),C_TBB)
  End If

   ' 目次を書き出さないスライド番号(1-based)
  Dim nSids()  As Integer
  Call bufRead_nList(hBufs(NOHI_MARK),nSids(),C_CMM) ' vbTab -> ","

  ' 書きだす場所
  Dim xPos0  As Integer
  Dim wid0   As Integer
  Dim Fsize0 As Integer
  Dim sizFlg As Integer
  sizFlg = 0
  If Not IsNull(hBufs(IFHI_MARK)) And Not hBufs(IFHI_MARK) = "" Then
    Call bufRead_siz(hBufs(IFHI_MARK),xPos0,wid0,Fsize0)
    sizFlg = 1
  End If
  'MsgBox("xPos0=" & xPos0 & vbNewLine & "wid0="  & wid0 & vbNewLine & "Fsize0=" & Fsize0)

  Dim xPos  As Integer
  Dim yPos  As Integer
  Dim wid   As Integer
  Dim hei   As Integer
  Dim Fsize As Integer

  'xPos  = 590
  'xPos = 5  ' 階層のときは左上
  'yPos  = 5
  'yPos = 1
  'wid   = 125
  'wid   = 200
  hei   = 20    ' 関係ない
  'Fsize = 10
  'Fsize = 8

  xPos = C_XPOSL
  yPos = C_YPOSL
  wid  = C_WIDEL
  Fsize= C_FSIZE

  ' 書き出す場所, 幅, フォントサイズの変更
  Call siz_MsgBox(sizFlg,xPos,wid,Fsize,xPos0,wid0,Fsize0)

  ' スライドへ書き出し
  Call write_TOC(xPos,yPos,wid,hei,Fsize,nSids(),Conts(),Pagesc(),C_HIERARCHY,0)

End Sub

Public Sub removeListHierarchy()
  Dim SlideObj   As Slide
  Dim ShapeObj   As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides          ' 各スライド
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1   ' 各シェイプ
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoTextBox Then           ' TextBox である
        If ShapeObj.Name = C_HIERARCHY Then          ' 名前が thierarchy である.
          ShapeObj.Delete                          ' オブジェクトを消去する.
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub

' 学生用資料を作成します
Public Sub pv_0100_flipTextForStudent()
  flipTextForStudy(1)
End Sub

' 教員用資料を作成します
Public Sub pv_0200_watermarkForTeacher()
  flipTextForStudy(0)
End Sub

' 教員用資料の作成 sw =0 にしたとき
'   キーワードが透かしで表示されたスライド
'   学生用では削除してあるスライドに透かしをいれる
' 学生用資料の作成: sw = 1 にしたとき
'   キーワードが置換されたスライド
'   答が載っているスライドを削除する
Private Sub flipTextForStudy(student As Integer)

  ' このルーチンでは以下を自動で行う:
  ' 1. 学生用に, キーワードが置換(_ 等)されたスライドを作成する.
  ' 2. 演習問題の答えが載っているスライドを削除する
  '
  ' 開いている powerpoint ファイル (hoge.pptm) に対して,
  ' 1. 置換情報, 削除スライド番号が書かれたテキストファイル(default=hoge.txt) を読み込む.
  ' 2. 元の pptm(pptx) file(hoge.pptm)  -> hoge_for_student.pptm に保存しなおし
  ' 3. 文字列置換とスライド削除

  ' hoge.txt の中身:
  ' ::%%%::
  ' Entrez Gene    _____ ____     *
  ' SNP            ___            1,2,3-10,20
  ' db___          dbSNP          1,2,20
  ' ::@@@::
  ' 5,8,11
  ' hoge.txt の中身終わり
  '
  ' hoge.txt の format:
  ' ::%%%::
  ' (置換する文字列)\t\t\t\t(置換後の文字列)\t\t\t\t(置換するページ)
  '    注: 置換するページについて:
  '        *: 全部のページ(この場合は省略可)
  '        1,2,3: 1,2,3 ページ
  '        3-10:  3 ページから 10 ページまで
  ' ::@@@::
  ' (ページ数),(ページ数)-(ページ数) ... <- 削除するページ数
  ' または以下のように, 保存するページ数を書く
  ' ::===::
  ' (ページ数),(ページ数)-(ページ数) ... <- 残すページ数

  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  ' text file が存在していなければ終了する
  If Len(Dir$(sNotesFilePath)) = 0 Then
    MsgBox (sNoteFilePath & " is missing")
    Exit Sub
  End If

  ' 保存する学生用資料のファイル名
  Dim sPptFileName As String
  If student = 1 Then
    spptFileName = newFileName("_for_student.pptm","do you want to save the file as ",1)
  Else
    spptFileName = newFileName("_for_teacher.pptm","do you want to save the file as ",1)
  End If

  ' 別名で上書き保存
  ' ここからはこのファイルが Active となるので注意する.
  If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

  ' file reading
  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)

  Dim fString() As String ' 置換元
  Dim tString() As String ' 置換先
  Dim Pagesc()  As String ' 置き換えるページ数 例: Pagesc(1)=(1,2,3,5)
  Dim flag      As Integer
  flag = 0
  If Not IsNull(hBufs(FLIP_MARK)) And Not hBufs(FLIP_MARK) = "" Then
    Call bufRead_exList(hBufs(FLIP_MARK),fString(),tString(),Pagesc(),C_TBB)
    flag = 1
  End If

  Dim nSids() As Integer
  Dim aSids() As Integer
  If Not IsNull(hBufs(DELS_MARK)) And Not hBufs(DELS_MARK) = "" Then
    Call bufRead_nList(hBufs(DELS_MARK),nSids(),C_CMM)  ' vbTab -> ","
  Else
    Call bufRead_nList(hBufs(ALIV_MARK),aSids(),C_CMM)  ' vbTab -> ","
  End If

  Dim npage As Integer
  ActivePresentation.Slides(1).Select
  npage = ActivePresentation.Slides.Count

  ' fString が定義されていれば置換する
  ' 2014/04/09
  If Not Sgn(fString) = 0 Then
    Dim fnum As Integer
    fnum = UBound(fString)

    ' 置換するページ
    ' doFlip(置換ペアID(0-based), スライドID(0-based))=1 -> 置換する
    Dim vBuf()     As String
    Dim iBuf()     As String
    Dim doFlip()   As Integer

    If flag = 1 Then
      ' 動的二次元配列の初期化
      ReDim doFlip(fnum)
      Dim i As Integer
      For i = 1 To npage
        ReDim doFlip(fnum, i)
      Next i

      Dim j As Integer
      For j = 0 To fnum
        Dim Pages() As Integer
        'MsgBox("j=" & j & vbNewLine & "Pagesc(j)=" & Pagesc(j))
        'Pages = Split(Pagesc(j),",")
        Call splitInt(Pagesc(j),Pages(),C_CMM)
        Dim k As Variant
        For Each k In Pages
          doFlip(j,CInt(k))=1
        Next k
      Next j

      ' 文字列置換
      Dim oSlide   As Slide
      Dim oShape   As Shape
      Dim Gry
      Gry = RGB(170,170,170)
      k = 1
      For Each oSlide In ActivePresentation.Slides   ' 各スライド
        For Each oShape In oSlide.Shapes             ' 各シェイプ
          ' MsgBox (oShape.Name & " " & oShape.Type & "+" & oShape.AutoShapeType)
          For j = 0 To fnum
            If Not IsNull(doFlip(j, k)) And doFlip(j, k) = 1 Then
              If student = 1 Then
                ' 置き換える
                Call FindnRe(oShape, fString(j), tString(j))
              Else
                ' 置き換えずに透かしで書く
                Call FindnCol(oShape, fString(j), Gry)
              End If
            End If
          Next j
        Next oShape
        k = k + 1
      Next oSlide
    End If
    
  End If

  ' 20131014
  ' 教員用 slide(穴は空いているがスライドの削除はされていない)を作成.
  ' スライドを削除する前に別名で保存しておく.
  'Dim sPptFileName_For_Teacher As String
  'sPptFileName_For_Teacher = Mid$(ActivePresentation.Name, 1, InStr(ActivePresentation.Name, ".") - 1) & "_for_teacher.pptm"
  'ActivePresentation.SaveAs (sCurrentFolder & sPptFileName_For_Teacher)

  Dim Gry2
  Gry2 = RGB(200,200,200)
  If student = 1 Then
    ' 学生用で必要のないスライドの削除
    Call removeSlide(nSids(),aSids(),npage)
  Else
    Call watermarkSlide(nSids(),aSids(),npage, Gry2)
  End If

  ' 20131014
  ' 学生用に削除した後の slide を hoge_for_student.pptm に保存しておく.
  ' いちいち 保存しますか と聞かれるのが面倒なので
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

End Sub

Private Function newFileName(suffix As String, message As String, mflg As Integer, Optional Fname As String) As String
  If IsNull(Fname) Or Fname = "" Then
    Fname = ActivePresentation.Name
  End If
  Dim sPptFileName As String
  'sPptFileName = Mid$(ActivePresentation.Name, 1, InStr(ActivePresentation.Name, ".") - 1)
  'sPptFileName = Mid$(Fname, 1, InStr(Fname,".") - 1)
  sPptFileName = Mid$(Fname, 1, InStrRev(Fname,".") - 1)  ' 2014/04/11
  sPptFileName = sPptFileName & suffix
  If mflg = 0 Then
    ' do nothing 確認しないで進む
  Else
    ' 確認する
    If MsgBox(message & "'" & sPptFileName & "'?", vbYesNo) = vbNo Then
      sPptFileName = InputBox("input file name", "ファイル名を入力", sPptFileName)
    End If
  End If
  newFileName = sPptFileName
End Function

Private Sub removeSlide(nSids() As Integer, aSids() As Integer, npage As Integer)
  ' 20130611
  ' i 番目スライド(1-based) を
  ' 削除する   dSld(i) = 1
  ' 削除しない dSld(i) = 0
  Dim dSld() As Integer
  If Not Sgn(nSids) = 0 Then
    Call checkd_slde(nSids(), dSld(), npage)
  ElseIf Not Sgn(aSides) = 0 Then
    Call checkw_slde(aSids(), dSld(), npage)
  End If

  If Not Sgn(dSld) = 0 Then
    Dim i As Integer
    For i = npage To 1 Step -1 ' スライドを削除するとスライド番号が変わるので, 後ろから計算する.
      If dSld(i) = 1 Then      ' 削除
        ActivePresentation.Slides(i).Delete
      End If
    Next i
  End If
End Sub

Private Sub watermarkSlide(nSids() As Integer, aSids() As Integer, npage As Integer, rgb As Variant)
  ' 20141114 from removeSlide
  ' i 番目スライド(1-based) に
  ' 印をつける  dSld(i) = 1
  ' つけない    dSld(i) = 0
  Dim dSld() As Integer
  If Not Sgn(nSids) = 0 Then
    Call checkd_slde(nSids(), dSld(), npage)
  ElseIf Not Sgn(aSides) = 0 Then
    Call checkw_slde(aSids(), dSld(), npage)
  End If

  If Not Sgn(dSld) = 0 Then
    Dim i As Integer
    For i = npage To 1 Step -1 ' 削除はしないので番号は変わらないが, removeSlide を踏襲する
      If dSld(i) = 1 Then      ' 印をつける
        With ActivePresentation.Slides.Item(i).Shapes.AddLine(C_XPOSL,C_YPOSL,C_XPOSB,C_YPOSB)
          .Name               = "watermarks"
          .Line.ForeColor.RGB= rgb
          '.Line.DashStyle     =msoLineRoundDot
          .Line.Style = msoLineSingle
          .Line.Weight        =C_LWID
          .ZOrder(msoSendToBack) ' 最下層に置く
        End With
      End If
    Next i
  End If
End Sub

Private Sub bufWrit_exList(prnt As String, sBuf As Variant, delm As String, offset() As Integer, Optional pExist As Variant)

  If Sgn(offset) = 0 Then
    If Not IsNull(sBuf) And Not sBuf = "" Then
      prnt = prnt & sBuf & vbNewLine
    End If
    Exit Sub
  End If

  Dim   fString() As String
  Dim   tString() As String
  Dim   Pagesc()  As String
  Call bufRead_exList(sBuf,fString(),tString(),Pagesc(),delm)

  Dim m As Integer
  m = UBound(fString)

  Dim k  As Integer
  For k = 0 To m
    Dim sForm As String
    sForm = "" 'Why ?
    Call sList2OFFSETsForm(Pagesc(k),sForm,C_CMM,offset(),pExist)
    If sForm <> "" Then  ' 2014/04/29 ページ数が無い場合の置換文字列は書かない
      prnt = prnt & fString(k) & delm & tString(k) & delm & sForm & vbNewLine
    End If
  Next k

End Sub

Private Sub bufRead_exList(sBuf As Variant, fString() As String, tString() As String, Pagesc() As String,delm As String)
  Dim aBuf() As String

  aBuf = Split(sBuf,vbNewLine) ' 一行ずつ分割

  Dim m As Integer
  m = UBound(aBuf) ' m = 行数 -1 (0-based)
  ReDim Preserve fString(m)
  ReDim Preserve tString(m)
  ReDim Preserve Pagesc(m)
  Dim sForm() As String
  ReDim Preserve sForm(m)

  Dim i As Integer
  For i = 0 To m
    Dim buf() As String
    buf=Split(aBuf(i),delm)
    Dim k As Integer
    k = 0
    Dim j As Integer
    For j = 0 To UBound(buf)
      If Not IsNull(buf(j)) And Not buf(j) = "" Then
        If k = 0 Then      ' from (置換元文字列)
          fString(i) = buf(j)
        ElseIf k = 1 Then ' to (置換先文字列)
          tString(i) = buf(j)
        ElseIf k = 2 Then ' page (置換するページのリスト)
          sForm(i)  = buf(j) ' とりあえず文字列として取り出す(後で処理する)
        End If
        k = k + 1
      End If
    Next
  Next i

  ' Pagesc
  Dim defList() As Integer
  Call mkDefList(defList())
  For i = 0 To m
    Call sForm2sList(sForm(i),Pagesc(i),defList(),C_CMM)
  Next i

End Sub


'
' page の表現の仕方
'
' sForm   = "1-3,4,6,10-12"    'text file
' iList() = (1,2,3,4,6,10,12)  ' 計算機内部
' sList   = "1,2,3,4,6,10,12"  ' 計算機内部(中間体)
' (defList() = (1,2,3, .... npage))
' の間の相互変換 script
Private Sub sForm2iList(sForm As String, iList() As Integer, delm As String, defList() As Integer)
  Dim vBuf() As String
  vBuf = Split(sForm,delm)
  Dim j As Integer
  j = 0
  Dim p As Integer
  For p = 0 To UBound(vBuf)
    If vBuf(p) = "*" Then     '全部
      Dim i As Integer
      For i = 0 To UBound(defList)
        ReDim Preserve iList(i+j)
        iList(i+j)=defList(i)
      Next i
      Exit For
    ElseIf IsNumeric(vBuf(p)) Then
      ReDim Preserve iList(j)
      iList(j) = CInt(vBuf(p))
      j = j + 1
    Else
      Dim iBuf() As String
      iBuf = Split(vBuf(p),"-")
      If IsNumeric(iBuf(0)) And Not IsNull(iBuf(1)) And IsNumeric(iBuf(1)) Then
        For i = CInt(iBuf(0)) To CInt(iBuf(1))
          ReDim Preserve iList(j)
          iList(j)=i
          j = j + 1
        Next i
      End If
    End If
  Next p
End Sub

Private Sub iList2sForm(iList() As Integer, sForm As String, delm As String)
  Dim j  As Integer
  Dim st As Integer
  Dim bf As Integer
  Dim stt As Integer
  stt = C_EXC
  For j = 0 To UBound(iList)
    st = iList(j)
    bf = iList(j)-1
    If Not st = C_EXC Then
      stt = j
      Exit For
    End If
  Next j
  If stt = C_EXC Then
    Exit Sub
  End If

  For j = stt To UBound(iList)
    Dim p As Integer
    p = iList(j)
    If p = C_EXC Then
      GoTo SKIP
    End If
    If Not p = bf + 1 Then
      If st < bf Then
        If st + 1 = bf Then
          sForm = sForm & CStr(st) & delm & CStr(bf) & delm
        Else
          sForm = sForm & CStr(st) & "-" & CStr(bf) & delm
        End If
      Else
        sForm = sForm & CStr(st) & delm
      End If
      st = p
    End If
    bf = p
SKIP:
  Next j

  If st < bf Then
    sForm = sForm & CStr(st) & "-" & CStr(bf)
  Else
    sForm = sForm & CStr(st)
  End If
End Sub

Private Sub iList2sList(iList() As Integer,sList As String, delm As String)
  ' alias
  Call joinInt(sList,iList(),delm)
End Sub

Private Sub sList2iList(sList As String, iList() As Integer, delm As String)
  ' alias
  Call splitInt(sList,iList(),delm)
End Sub

Private Sub sForm2sList (sForm As String, sList As String,defList() As Integer,delm As String)
  Dim iList() As Integer
  Call sForm2iList(sForm,iList(),delm,defList())
  Call iList2sList(iList(),sList,delm)
End Sub

Private Sub sList2sForm(sList As String, sForm As String, delm As String)
  Dim iList() As Integer
  Call sList2iList(sList,iList(),delm)
  Call iList2sForm(iList(),sForm,delm)
End Sub

Private Sub sList2OFFSETsForm(sList As String, sForm As String, delm As String, offset() As Integer, Optional pExist As Variant)
  Dim iList() As Integer
  Call sList2iList(sList,iList(),delm)
  Call iList2OFFSETsForm(iList(),sForm,delm,offset(),pExist)
End Sub

Private Sub iList2OFFSETsForm(iList() As Integer, sForm As String, delm As String,offset() As Integer, Optional pExist As Variant)
  Dim j As Integer
  For j = 0 To UBound(iList)
    'If pExist Is Nothing Then
    If IsMissing(pExist) Then
        iList(j) = iList(j) + oFFst(iList(j),offset())
    Else
      If Not iList(j) > UBound(pExist) Then
        If pExist(iList(j)) = 1 Then
          iList(j) = iList(j) + oFFst(iList(j),offset())
        Else
          iList(j) = C_EXC
        End If
      Else
        iList(j) = C_EXC
      End If
    End If
  Next j
  Call iList2sForm(iList(),sForm,delm)
End Sub

Private Sub joinInt(Pagesc As String, iArry() As Integer, delm As String)
  Pagesc = ""
  Dim i As Integer
  For i = 0 To UBound(iArry)-1
    If Not iArry(i) = C_EXC Then
      Pagesc = Pagesc & CStr(iArry(i)) & delm
    End If
  Next i
  Pagesc = Pagesc & CStr(iArry(UBound(iArry)))
End Sub

Private Sub splitInt(Pagesc As String, iArry() As Integer, delm As String)
  Dim aBuf() As String
  aBuf = Split(Pagesc,delm)
  Dim m As Integer
  m = UBound(aBuf)
  ReDim iArry(m)
  Dim i As Integer
  For i = 0 To m
    iArry(i)=CInt(aBuf(i))
  Next i
End Sub

Private Sub mkDefList(defList() As Integer, Optional pPt As Presentation)

  Dim npage As Integer
  ' スライドの枚数を得る
  If pPt Is Nothing Then
    ActivePresentation.Slides(1).Select
    npage = ActivePresentation.Slides.Count
  Else
    npage = pPt.Slides.Count
  End If

  ReDim defList(npage-1)
  Dim i As Integer
  For i = 0 To npage-1
    defList(i)=i+1 ' 全部使う場合の Pages list (1,2,3,....npages)
  Next i
End Sub

'
' flip2
'
Private Sub FindnCol(oShape As Shape, fString As String, rgb As Variant)
  Dim i       As Integer
  ' shape のタイプによって分類
  'On Error Resume Next
  Select Case oShape.Type
    Case msoTable
      Call ColorTable(oShape,fString,rgb)
    Case msoGroup
      For i = 1 To oShape.GroupItems.Count
         Call FindnCol(oShape.GroupItems(i), fString, rgb)
      Next
    Case msoDiagram
      For i = 1 To oShape.Diagram.Nodes.Count
        Call FindnCol(oShape.Diagram.Nodes(i).TextShape, fString, rgb)
      Next
    Case msoSmartArt
      For i = 1 To oShape.SmartArt.AllNodes.Count
        Call ColorText2(oShape,i,fString, rgb)
      Next
    Case Else
      If oShape.HasTextFrame Then
        If oShape.TextFrame.HasText Then
          Call ColorText(oShape,fString, rgb)
        End If
      ElseIf oShape.HasSmartArt Then   ' placeholder 内に SmartArt がある場合 等.
        For i = 1 To oShape.SmartArt.AllNodes.Count
          Call ColorText2(oShape,i,fString,rgb)
        Next
      End If
  End Select
End Sub

Private Sub ColorTable(oShape As Shape, fString As String, rgb As Variant)
  Dim oTxtRng
  Dim oTmpRng
  Dim iRows As Integer
  Dim iCols As Integer
  For iRows = 1 To oShape.Table.Rows.Count
    For iCols = 1 To oShape.Table.Rows(iRows).Cells.Count
      Set oTxtRng = oShape.Table.Rows(iRows).Cells(iCols).Shape.TextFrame2.TextRange ' TextFrame2 を使った.
      Set oTmpRng = oTxtRng.Find(fString)
      Do While ((Not oTmpRng = "") and (Not oTmpRng Is Nothing))
        oTmpRng.Font.Fill.ForeColor.RGB=rgb ' 色の指定の仕方も変わっている
        oTmpRng.Font.Italic    = msoTrue
        oTmpRng.Font.UnderlineStyle=msoUnderlineWavyHeavyLine ' 波線を引いている
        Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
      Loop
    Next
  Next
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

Private Sub ColorText(oShape As Shape, fString As String, rgb As Variant)
  Dim oTxtRng
  Dim oTmpRng
  Set oTxtRng = oShape.TextFrame2.TextRange  ' TextFrame2 を使った.
  Set oTmpRng = oTxtRng.Find(fString)
  'MsgBox oTmpRng
  ' TextFrame2 にすると oTmpRng が "" になることがあって,
  ' Nothing の条件だけだとそのせいでエラーとなることがある.
  ' よくわかんないので条件を複数付けておく.
  ' Activpresentation.Slides(1).Shapes(1).TextFrame2.TextRange.Font.
  Do While ((Not oTmpRng = "") and (Not oTmpRng Is Nothing))
    oTmpRng.Font.Fill.ForeColor.RGB=rgb ' 色の指定の仕方も変わっている. 面倒くさい.
    oTmpRng.Font.Italic    = msoTrue
    oTmpRng.Font.UnderlineStyle=msoUnderlineWavyHeavyLine ' 波線を引いている
    Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
  Loop
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

Private Sub ColorText2(oShape As Shape, ByVal i As Integer, fString As String, rgb As Variant)
  Dim oTxtRng
  Dim oTmpRng
  Set oTxtRng = oShape.SmartArt.AllNodes(i).TextFrame2.TextRange ' これは元々 TextFrame2 だった.
  Set oTmpRng = oTxtRng.Find(fString)
  Do While ((Not oTmpRng = "") and (Not oTmpRng Is Nothing)) ' Nothing だけでエラーになったこと無いけど念のため
    oTmpRng.Font.Fill.ForeColor.RGB=rgb ' 色の指定の仕方も変わっている.
    oTmpRng.Font.Italic    = msoTrue
    oTmpRng.Font.UnderlineStyle=msoUnderlineWavyHeavyLine ' 波線を引いている
    Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
  Loop
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

'
' 下線のみを引く場合(TextFrame2 を使わなくて良い. より安全)
'
'Private Sub ColorTable(oShape As Shape, fString As String, rgb As Variant)
'  Dim oTxtRng
'  Dim oTmpRng
'  Dim iRows As Integer
'  Dim iCols As Integer
'  For iRows = 1 To oShape.Table.Rows.Count
'     For iCols = 1 To oShape.Table.Rows(iRows).Cells.Count
'       Set oTxtRng = oShape.Table.Rows(iRows).Cells(iCols).Shape.TextFrame.TextRange
'       Set oTmpRng = oTxtRng.Find(fString)
'       Do While Not oTmpRng Is Nothing
'         oTmpRng.Font.Color.RGB = rgb
'         oTmpRng.Font.Italic    = msoTrue
'         oTmpRng.Font.Underline = msoTrue  ' 下線を引く
'         Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
'       Loop
'     Next
'   Next
'   Set oTxtRng = Nothing
'   Set oTmpRng = Nothing
' End Sub
'
'End Sub

' Private Sub ColorText(oShape As Shape, fString As String, rgb As Variant)
'   Dim oTxtRng
'   Dim oTmpRng
'   Set oTxtRng = oShape.TextFrame.TextRange
'   Set oTxtRng = oShape.TextFrame2.TextRange
'   Set oTmpRng = oTxtRng.Find(fString)
'   MsgBox oTmpRng
'   Do While Not oTmpRng Is Nothing
'     oTmpRng.Font.Color.RGB = rgb
'     oTmpRng.Font.Italic    = msoTrue
'     oTmpRng.Font.Underline = msoTrue
'     oTmpRng.Font.UnderlineStyle=msoUnderlineWavyLine
'     Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
'   Loop
'   Set oTxtRng = Nothing
'   Set oTmpRng = Nothing
' End Sub

' Private Sub ColorText2(oShape As Shape, ByVal i As Integer, fString As String, rgb As Variant)
'   Dim oTxtRng
'   Dim oTmpRng
'   Set oTxtRng = oShape.SmartArt.AllNodes(i).TextFrame2.TextRange
'   Set oTmpRng = oTxtRng.Find(fString)
'   Do While Not oTmpRng Is Nothing
'     oTmpRng.Font.Color.RGB = rgb
'     oTmpRng.Font.Italic    = msoTrue
'     oTmpRng.Font.Underline = msoTrue
'     Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
'   Loop
'   Set oTxtRng = Nothing
'   Set oTmpRng = Nothing
' End Sub

'
' flip 関連
'
Private Sub FindnRe(oShape As Shape, fString As String, tString As String)
  Dim i       As Integer
  ' shape のタイプによって分類
  'On Error Resume Next
  Select Case oShape.Type
    Case msoTable
      Call ReplaceTable(oShape,fString,tString)
    Case msoGroup
      For i = 1 To oShape.GroupItems.Count
         Call FindnRe(oShape.GroupItems(i), fString, tString)
      Next
    Case msoDiagram
      For i = 1 To oShape.Diagram.Nodes.Count
        Call FindnRe(oShape.Diagram.Nodes(i).TextShape, fString, tString)
      Next
    Case msoSmartArt
      For i = 1 To oShape.SmartArt.AllNodes.Count
        Call ReplaceText2(oShape,i,fString,tString)
      Next
    Case Else
      If oShape.HasTextFrame Then
        If oShape.TextFrame.HasText Then
          Call ReplaceText(oShape,fString,tString)
        End If
      ElseIf oShape.HasSmartArt Then   ' placeholder 内に SmartArt がある場合 等.
        For i = 1 To oShape.SmartArt.AllNodes.Count
          Call ReplaceText2(oShape,i,fString, tString)
        Next
      End If
  End Select
  ' メモリ開放
End Sub

Private Sub ReplaceTable(oShape As Shape, fString As String, tString As String)
  Dim oTxtRng
  Dim oTmpRng
  Dim iRows As Integer
  Dim iCols As Integer
  For iRows = 1 To oShape.Table.Rows.Count
    For iCols = 1 To oShape.Table.Rows(iRows).Cells.Count
      Set oTxtRng = oShape.Table.Rows(iRows).Cells(iCols).Shape.TextFrame.TextRange
      Set oTmpRng = oTxtRng.Replace(fString, tString)
      Do While Not oTmpRng Is Nothing
        'oTmpRng.Select
        Set oTmpRng = oTxtRng.Replace(fString, tString, After:=oTmpRng.Start + oTmpRng.Length)
      Loop
    Next
  Next
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

Private Sub ReplaceText(oShape As Shape, fString As String, tString As String)
  Dim oTxtRng
  Dim oTmpRng
  Set oTxtRng = oShape.TextFrame.TextRange
  'oTxtRng.Select
  Set oTmpRng = oTxtRng.Replace(fString, tString)
  Do While Not oTmpRng Is Nothing
    Set oTmpRng = oTxtRng.Replace(fString, tString, After:=oTmpRng.Start + oTmpRng.Length)
  Loop
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

Private Sub ReplaceText2(oShape As Shape, ByVal i As Integer, fString As String, tString As String)
  Dim oTxtRng
  Dim oTmpRng
  Set oTxtRng = oShape.SmartArt.AllNodes(i).TextFrame2.TextRange
  'MsgBox(oShape.SmartArt.AllNodes(1).TextFrame2.TextRange.Text & "; " & oShape.SmartArt.AllNodes.Count)
  Set oTmpRng = oTxtRng.Replace(fString,tString)
  Do While Not oTmpRng Is Nothing
    Set oTmpRng = oTxtRng.Replace(fString, tString, After:=oTmpRng.Start + oTmpRng.Length)
  Loop
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

' for debug (SmartArt)
Private Sub findSmartArt(oDummy As Shape)
  Dim oSLide As Slide
  Dim oShape As Shape
  For Each oSlide In ActivePresentation.Slides   ' 各スライド
    For Each oShape In oSlide.Shapes             ' 各シェイプ
      If oShape.Type = msoSmartArt Then
        'If oShape.GroupItems.Count And oShape.GroupItems(1).HasTextFrame Then
        '  MsgBox(oShape.GroupItems(1).TextFrame.TextRange.Text & ": " & oShape.GroupItems.Count)
        'End If
        MsgBox("allnodes: " & oShape.SmartArt.AllNodes.Count)
      End If
      If oShape.Type = msoPlaceholder Then
        If oShape.HasSmartArt Then
          If oShape.GroupItems.Count And oShape.GroupItems(1).HasTextFrame Then
            MsgBox("PlaceHolder: " & oShape.GroupItems(1).TextFrame.TextRange.Text & ": " & oShape.GroupItems.Count)
          End If
        End If
      End If
    Next oShape
  Next oSlide
  ' メモリ解放
  Set oSlide = Nothing
  Set oShape = Nothing
End Sub

Private Sub findSmartArt2(oDummy As Shape)
  Dim oSLide As Slide
  Dim oShape As Shape
  For Each oSlide In ActivePresentation.Slides   ' 各スライド
    For Each oShape In oSlide.Shapes             ' 各シェイプ
      If oShape.Type = msoSmartArt Then
        'If oShape.GroupItems.Count And oShape.GroupItems(1).HasTextFrame Then
        '  MsgBox(oShape.GroupItems(1).TextFrame.TextRange.Text & ": " & oShape.GroupItems.Count)
        'End If
        MsgBox("allnodes: " & oShape.SmartArt.AllNodes(1).TextFrame2.TextRange.Text & "; " & oShape.SmartArt.AllNodes.Count)
      End If
      If oShape.Type = msoPlaceholder Then
        If oShape.HasSmartArt Then
          'If oShape.GroupItems.Count And oShape.GroupItems(1).HasTextFrame Then
           ' MsgBox("PlaceHolder: " & oShape.GroupItems(1).TextFrame.TextRange.Text & ": " & oShape.GroupItems.Count)
          'End If
          MsgBox("placeholder: " & oShape.SmartArt.AllNodes(1).TextFrame2.TextRange.Text & "; " & oShape.SmartArt.AllNodes.Count)
        End If
      End If
    Next oShape
  Next oSlide
  ' メモリ解放
  Set oSlide = Nothing
  Set oShape = Nothing
End Sub

' Sub FindnRe(oShape As Shape, fString As String, tString As String)
'   Dim oTxtRng As TextRange
'   Dim oTmpRng As TextRange
'   Dim i       As Integer
'   Dim iRows   As Integer
'   Dim iCols   As Integer
'   'Dim oShpTmp As Shape
'
'   ' shape のタイプによって分類
'   'On Error Resume Next
'   Select Case oShape.Type
'     Case msoTable
'       For iRows = 1 To oShape.Table.Rows.Count
'         For iCols = 1 To oShape.Table.Rows(iRows).Cells.Count
'           Set oTxtRng = oShape.Table.Rows(iRows).Cells(iCols).Shape.TextFrame.TextRange
'           Set oTmpRng = oTxtRng.Replace(fString, tString)
'           Do While Not oTmpRng Is Nothing
'             'oTmpRng.Select
'             Set oTmpRng = oTxtRng.Replace(fString, tString, After:=oTmpRng.Start + oTmpRng.Length)
'           Loop
'         Next
'       Next
'     Case msoGroup
'       For i = 1 To oShape.GroupItems.Count
'         Call FindnRe(oShape.GroupItems(i), fString, tString)
'       Next
'     Case msoDiagram
'       For i = 1 To oShape.Diagram.Nodes.Count
'         Call FindnRe(oShape.Diagram.Nodes(i).TextShape, fString, tString)
'       Next
'     Case msoSmartArt
'       For i = 1 To oShape.GroupItems.Count
'         Call FindnRe(oShape.GroupItems(i), fString, tString)
'       Next
'     Case Else
'       If oShape.HasTextFrame Then
'         If oShape.TextFrame.HasText Then
'           Set oTxtRng = oShape.TextFrame.TextRange
'           'oTextRng.Select
'           Set oTmpRng = oTxtRng.Replace(fString, tString)
'           Do While Not oTmpRng Is Nothing
'             Set oTmpRng = oTxtRng.Replace(fString, tString, After:=oTmpRng.Start + oTmpRng.Length)
'           Loop
'         End If
'       End If
'   End Select
' End Sub


'
' old version
'
'Public Sub pageNum()
Private Sub pageNum()

  ' 場所指定
  Dim xPos, yPos, wid, hei
  xPos = 686
  yPos = 527

  wid = 40
  hei = 28.875

  ' font
  Dim Fname, Ename, Fsize
  Fname = "Arial"
  Ename = "ＭＳ Ｐゴシック"
  Fsize = 10

  ActivePresentation.Slides(1).Select     ' 一枚目のスライドを選択
  Dim npage
  npage = ActivePresentation.Slides.Count ' スライドの総数を数える

  Dim i As Integer
  Dim Shp
  For i = 1 To npage
    ActivePresentation.Slides(i).Select
   ' i 枚目のスライド
    Set Shp = ActivePresentation.Slides.Item(i).Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, wid, hei)
    Shp.Select
    ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=i, Length:=0).Select
    With ActiveWindow.Selection.TextRange
      .Text = i & "/" & npage
      With .Font
           .NameAscii = Fname
           .NameFarEast = Ename
           .NameOther = Fname
           .Size = Fsize
           .Bold = msoFalse
           .Italic = msoFalse
           .Underline = msoFalse
           .Shadow = msoFalse
           .Emboss = msoFalse
           .BaselineOffset = 0
           .AutoRotateNumbers = msoTrue
           .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0)
      End With
    End With
    Shp.Name = C_PAGENUM  ' textbox に名前をつけておく
  Next
End Sub

'Public Sub pageCountsBox()
Private Sub pageCountsBox()
  ' 色設定
  Dim Red, Blk, Gry, Wht, Col
  Red = RGB(255, 0, 0)
  Blk = RGB(0, 0, 0)
  Gry = RGB(70, 70, 70)
  Wht = RGB(255, 255, 255)
  Col = RGB(0, 150, 0)
  'Red=RED
  'Blk=BLACK
  'Gry=GRAY2
  'Wht=WHITE
  'Col=GREEN

  ' 場所設定
  Dim lef, top, wid, hei, stp, wei
  lef = -9
  top = 532
  wid = 8
  hei = 8
  stp = 8.5
  wei = 0.5

  ActivePresentation.Slides(1).Select
  Dim npage As Integer
  npage = ActivePresentation.Slides.Count
  Dim ShapeList() As String
  ReDim ShapeList(npage)

  Dim i As Integer
  Dim Shp
  For i = 1 To npage
    'ActivePresentation.Slides(i).Select
    Dim j As Integer
    For j = 1 To npage
      Set Shp = ActivePresentation.Slides.Item(i).Shapes.AddShape(msoShapeRectangle, lef + (j * stp), top, wid, hei)
      'Shp.Select
      '2012/04/24 initialize fill pattern
      Set oSp=ActivePresentation.Slides(i).Shapes(ActivePresentation.Slides(i).Shapes.Count)
      oSp.Fill.Solid
      oSp.Line.Weight = wei
     ' 枠線
      If (j Mod 10 = 0) Then
        oSp.Line.ForeColor.RGB=Red
      Else
        oSp.Line.ForeColor.RGB=Blk
      End If
     ' 塗りつぶし
      If (j < i) Then
        oSp.Fill.ForeColor.RGB=Gry
      ElseIf (j = i) Then
        oSp.Fill.ForeColor.RGB=Col
      Else
        oSp.Fill.ForeColor.RGB=Wht
      End If
      ShapeList(j) = Shp.Name
    Next
    ' グループ化して名前をつける
    ActivePresentation.Slides(i).Shapes.Range(ShapeList()).Group.Name = C_COUNTBOX
  Next
  ActivePresentation.Slides(1).Select
  ' メモリ解放
End Sub

Public Sub source_copy()
  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)
  
  Set hBufs = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)
  ' java file 名の取得
  Dim jFiles() As String 
  jFiles= Split(hBufs(SOUR_MARK),vbNewLine)
  
  Dim i As Integer
  For i= 0 To UBound(jFiles)
    Dim page As Integer
    page = i + 1
    Dim sContent As String
    Call fileContents(filePath(jFiles(i)),sContent)   ' ファイルの中身を得る
    ActivePresentation.Slides.Add page, ppLayoutBlank ' 白紙スライドの追加
    ActivePresentation.Slides(page).Select
    ' 追加したスライドに textbox を作成
    Set Shp=ActivePresentation.Slides.Item(page).Shapes.AddTextbox(msoTextOrientationHorizontal,C_XPOSS,C_YPOSS,C_XWIDS,C_YWIDS)
    Shp.Select
    ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=page, Length:=0).Select
    With ActiveWindow.Selection.TextRange
      .Text = sContent ' ファイルの中身を書き込む
      With .Font
        .NameAscii = C_SNAME
        .NameFarEast = C_SNAME
        .NameOther = C_SNAME
        .Size = C_SSIZE
        .Bold = msoFalse
        .Italic = msoFalse
        .Underline = msoFalse
        .Shadow = msoFalse
        .Emboss = msoFalse
        .BaselineOffset = 0
        .AutoRotateNumbers = msoTrue
        .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0)
      End With
    End With
    'Shp.Name = C_PAGENUM  ' textbox に名前をつけておく
  Next i
End Sub

Public Sub write_txt_down_hierarchy()
  '
  ' ::|||:: の修正
  '
  ' 全体の hierarchy を一つ下げる
  '

  Dim npage As Integer
  npage = ActivePresentation.Slides.Count

  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)

  Dim aBuf() As String
  aBuf()=Split(hBufs(GTOC_MARK),vbNewLine)
  
  Dim m As Integer
  m = UBound(aBuf)

  Dim hStr  As String
  Dim page  As Integer
  Dim buf() As String
  page = npage
  For j=m To 0 Step -1
    Dim s As String
    s = Left$(aBuf(j),1)
    If s = "" OR s = " " OR s = vbTab Then
      hStr = aBuf(j) + vbNewLine + hStr ' 何もしない
    ElseIf s = "*" Then
      hStr = "*" + aBuf(j) + vbNewLine + hStr ' * があるときは一つ足すだけ
      buf() = Split(aBuf(j),vbTab)
      page  = Val(buf(UBound(buf)))
    Else
      hStr = "*" + vbTab + aBuf(j) + vbTab + Trim(Str(page)) + vbNewLine + hStr' 無いときは *\t をつける 
    End If
  Next j

  ' GTOC_MARK key のデータを消去
  hBufs.Remove(GTOC_MARK)
  ' GTOC_MARK key のデータを新しく作成
  hBufs.Add GTOC_MARK, hStr

  ' 上書きする点にちゅうい
  If MsgBox("overwrite hierarchy to " &  sNotesFilePath & ": OK ?",vbYesNo) = vbYes Then
    Call hashPrint(hBufs,sNotesFilePath)
  Else
    MsgBox("do nothing")
  End If

End Sub

Public Sub write_txt_ListContents()

  '
  ' 階層目次 ::|||:: GTOC_MARK から,
  ' ListContents 用の目次 ::###:: (右上に書く内容)
  ' のリストを自動でとってくる
  ' ListContents で書かせる階層を指定する
  '
  
  Dim npage As Integer
  npage = ActivePresentation.Slides.Count
  
  Dim maxH As Integer
  maxH = C_MAXH
  If MsgBox("the hire num of ListContents TOC: " & maxH & " ?",vbYesNo) = vbNo Then
    maxH = InputBox("the hierarchy number: ",maxH)
  End If
  
  ' text file(path) の取得
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  Set hBufs   = New Dictionary ' 読み込んだ文字列
  Call fileRead2(sNotesFilePath,hBufs)
  
  Dim aBuf() As String
  aBuf()=Split(hBufs(GTOC_MARK),vbNewLine)
  
  Dim m As Integer
  m = UBound(aBuf)
  
  Dim j      As Integer
  Dim pnum() As Integer
  ReDim pnum(maxH)
  For j=0 To maxH
    pnum(j)=npage   ' 階層 1 のページ = pnum(1)
  Next j

  Dim hStr   As String
  Dim buf()  As String
  Dim star   As Integer
  For j=m To 0 Step -1  ' 最後から見ていく.
    If Left$(aBuf(j),1) = "*" Then
      buf()=Split(aBuf(j),C_TBB)
      star = Len(buf(0))
      If Not star > maxH Then
        If Not buf(UBound(buf)) = Trim(Str(pnum(star))) Then
          hStr = buf(1) + vbTab + buf(UBound(buf)) + "-" + Trim(Str(pnum(star))) + vbNewLine + hStr
        Else
          hStr = buf(1) + vbTab + buf(UBound(buf)) + vbNewLine + hStr ' 1 page だけの場合
        End If
        pnum(star)=Val(buf(UBound(buf)))-1
      End If
    End If
  Next j

  ' 上書きする点にちゅうい
  If MsgBox("add ListContents to " &  sNotesFilePath & ": OK ?",vbYesNo) = vbYes Then
    ' TOCS_MARK key のデータを消去
    If hBufs.Exists(TOCS_MARK) Then
      hBufs.Remove(TOCS_MARK)
    End If
    ' TOCS_MARK key のデータを新しく作成
    hBufs.Add TOCS_MARK, hStr
    Call hashPrint(hBufs,sNotesFilePath)
  Else
    MsgBox("do nothing")
  End If
  
End Sub

' 文字化けを心配しなくて良いとき
'Private Sub fileContents(sNotesFilePath As String, sContent As String)
'  If Dir(sNotesFilePath) = "" Then  ' ファイルの存在を確認
'    Exit Sub
'  End If
'
'  Dim iNotesFileNum As Integer
'  iNotesFileNum = FreeFile()
'  Open sNotesFilePath For Input As iNotesFileNum
'
'  Dim sBuf As String
'  sContent=""
'  Do Until EOF(iNotesFileNum)
'    Line Input #iNotesFileNum,sBuf
'    sContent = sContent & sBuf & vbNewLine ' 書いてある文字列を入れておく
'  Loop
'  Close iNotesFileNum
'End Sub

' UTF-8 テキストファイルを読み込むとき
Private Sub fileContents(sNotesFilePath As String, sContent As String)
  If Dir(sNotesFilePath) = "" Then  ' ファイルの存在を確認
    Exit Sub
  End If

  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    .LoadFromFile sNotesFilePath
    sContent = .ReadText
    .Close
  End With
  
End Sub

Private Function filePath(Fname As String) As String
  Fname=Replace(Fname,"/",Application.PathSeparator)
End Function

End Function
Private Function filePath_org(Fname As String) As String
  Dim Op As Variant
  Op = Application.OperatingSystem
  If Op Like "Macintosh*" Then
    Fname = Replace(Fname,"\",":")
    If InStr(Fname, ":") = 0 Then
      Fname = ActivePresentation.Path & ":" & Fname
    ElseIf Left$(Fname,2) = ".:" Then
      Fname = ActivePresentation.Path & ":" & Mid(Fname,2)
    End If
  Else
    ' もしFname がファイル名だけの場合は, カレントディレクトリにあるとみなして
    ' 絶対パスにする.
    'あるいは "." から始まる場合には相対パスで書いてあると考えて
    ' 今いるディレクトリ名を追記して絶対パスにする
    ' それ以外の場合は, 絶対パスで書いてあるとかんがえる.
    If InStr(Fname, "\") = 0 Then
      Fname = ActivePresentation.Path & "\" & Fname
      'ElseIf Left$(Fname,2) = ".\" Then   ' 2014/04/11 削除
      '  Fname = ActivePresentation.Path & "\" & Mid(Fname,2)
    ElseIf Left(Fname,1) = "." Then ' 2014/04/11 よくわからない.
      Fname = ActivePresentation.Path & "\" & Fname
    End If
  End If
  filePath=Fname
End Function