Attribute VB_Name = "AutoPowerPointDecoration"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' text file ��ǂݍ��� format �̕ύX
' see Private Sub fireRead2
' 1. ���ڈ�ŏ��ƍŌ�� "::" ������
' 2. �ڈ�{�̂� 3 �����̕�����(���ł��ǂ�) ��: @@@
' 3. hBufs("@@@") �ŕ����񂪎擾�ł���
' 4. ����ꂽ������� parse ���Ēm�肽�����𓾂�
'
' ::@@@::
' (��������)
' ::$$$::
' (��������)

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
Private Const C_TITLE     As String  = "�ڎ�"
Private Const C_PLCHOLDER As String  = "Placeholder"
Private Const C_TITLE1    As String  = "Title"
Private Const C_FSIZE     As Integer = 8
Private Const C_XPOSR     As Integer = 610    ' pageListContents default X ���W
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


Private Const COMT_MARK   As String  = "'"    ' comment �s�̎w��
Private Const INFO_MARK   As String  = "::"   ' �����̖͂ڈ�
Private Const GTOC_MARK   As String  = "|||"  ' grayTOC �p�̖ڈ�           grayTOC
Private Const CPPT_MARK   As String  = "+++"  ' collectPpt �p�̖ڈ�        collectPpt
Private Const SKIP_MARK   As String  = "---"  ' ��������y�[�W             skipSlide
Private Const TOCS_MARK   As String  = "###"  ' �ڎ��ƃy�[�W��             pageListContents, pageListHierarchy
Private Const NOWT_MARK   As String  = "$$$"  ' �ڎ������������Ȃ��y�[�W   pageListContents
Private Const IFNT_MARK   As String  = "!!!"  ' �������� font ���         pageListContents
Private Const TOHI_MARK   As String  = "^^^"  ' �ڎ��ƃy�[�W��             pageListHierarchy
Private Const NOHI_MARK   As String  = "&&&"  ' �ڎ������������Ȃ��y�[�W   pageListHierarchy
Private Const IFHI_MARK   As String  = "???"  ' �������� font ���         pageListHierarchy
Private Const FLIP_MARK   As String  = "%%%"  ' �u�������񃊃X�g           zFlipTextForStudent
Private Const DELS_MARK   As String  = "@@@"  ' �폜�X���C�h               zFlipTextForStudent
Private Const ALIV_MARK   As String  = "==="  ' �ۑ��X���C�h(�폜 @@@�D��) zFlipTextForStudent
Private Const SOUR_MARK   As String  = "sss"  ' �t�@�C���̒��g�𒼐� ppt �ɏ������� source_copy

Private Const C_FNAME     As String  = "Arial"
Private Const C_ENAME     As String  ="�l�r �o�S�V�b�N"
Private Const C_SNAME     As String  ="�l�r �S�V�b�N"    ' source_copy

Private Const C_LWID      As Integer = 5      ' watermarkSlide ���̑���

'Private Const BLACK = RGB(0,0,0)
'Private Const GRAY  = RGB(150,150,150)
'Private Const RED   = RGB(255,0,0)
'Private Const GRAY2 = RGB(70,70,70)
'Private Const WHITE = RGB(255,255,255)
'Private Const GREEN = RGB(0,150,0)

' GFLAG=0 skip dialog
' GFLAG=1 dialog to confirm and change parameters

Public Sub p0_standard() ' skip dialog
  collectPpt(0)
  grayTOC(0)
  pageCountsBox2(0)
  pageNum2(0)
  pageListContents2(0)
  pageListHierarchy2(0)
End Sub

Public Sub p00_collectPpt()
  collectPpt(1)
End Sub

Public Sub p01_grayTOC()
  grayTOC(1)
End Sub

Public Sub p02_pageCountsBox()
  pageCountsBox2(1)
End Sub

Public Sub p03_pageNum()
  pageNum2(1)
End Sub

Public Sub p04_pageListContents()
  pageListContents2(1)
End Sub

Public Sub p05_pageListHierarchy()
  pageListHierarchy2(1)
End Sub

Public Sub p6_flipTextForStudent()
  Call flipTextForStudy(1,0)
End Sub

Public Sub p7_watermarkForLecture()
  Call flipTextForStudy(0,0)
End Sub

Public Sub p8_removeGrayTOC_asNewName()
  removeGrayTOC_asNewName(0)
End Sub


' text file �� �g�� powerpoint �t�@�C���Ƃ��̃y�[�W�����w�肵�Ă�����,
' ����𓮂������ƂŎw�肵���X���C�h������ powerpoint �t�@�C���Ɏ�荞�܂�܂�.
' NEW: text file �������, ��荞�񂾂��Ƃɂ��y�[�W�̂���(offset)���l������
' �V�����������ꂽ text file ���쐬���܂�.
Public Sub collectPpt(GFLAG As Integer)

  ' �ݒ�t�@�C����ǂݍ���,
  ' �����ɏ����ꂽ���� ppt �t�@�C���Ǝw�肵���y�[�W���R�s�[����
  ' �ŏ��̔ԍ� (0 or 1) ��, 0 �̂Ƃ� = �X�^�C���̓R�s�[���Ȃ�. 1 �̂Ƃ� = �X�^�C�����R�s�[)
  ' ::+++::
  ' 0   hoge.ppt    1-3,10,45,44
  ' 1   fuga.pptx   20,30,50-60
  ' 1   hoge.ppt    45
  ' [2018-02-06 Tue]

  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,GFLAG)

  '''Set hBufs = New Dictionary
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  ' �t�@�C�����̎擾
  ' Desgns(N)  �R�s�[�̕��@(0: �X�^�C���� target ppt 1: �X�^�C����ێ����ăR�s�[)
  ' Fnames(N)  �p���� ppt �t�@�C���̃��X�g
  ' Pdummy(N)  �y�[�W�������������X�g
  Dim   Desgns() As Integer
  Dim   Fnames() As String
  Dim   Pagesc() As String
  'vbTab �ŌŒ�("," �ɕύX�͂ł��Ȃ�. Pagesc ��"," ��؂�Ȃ̂�)
  Call bufRead_File(hBufs.Item(CPPT_MARK),Desgns(),Fnames(),Pagesc(),C_TBB)

  ' ���J���Ă���X���C�h
  Dim pTo  As Presentation
  Set PTo = Application.ActivePresentation

  ' �ۑ����鎑���̃t�@�C����
  Dim sPptFileName As String
  spptFileName = newFileName("_intg.pptm","do you want to save the file as ",GFLAG) ' [2018-02-06 Tue]

  ' �ʖ��ŏ㏑���ۑ�
  ' ��������͂��̃t�@�C���� Active �ƂȂ�̂Œ��ӂ���.
  ' [2018-02-06 Tue]
  Dim yvb As Integer
  If GFLAG = 1 Then
    yvb = MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo)
  Else
    yvb = vbYes
  End If
  'If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
  If yvb = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

  Dim m As Integer
  m = UBound(Fnames)

  Dim sttIns() As Integer
  ReDim sttIns(m)

  ' ppt �t�@�C���̎�荞��
  Dim k As Integer
  For k = 0 To m
    Dim pFr As Presentation ' ppt �t�@�C�����J���܂�.
    Set pFr = Presentations.Open(FileName:=Fnames(k),ReadOnly:=msoFalse)
    Dim sttIn As Integer
    'sttIn = ActivePresentation.Slides.Count + 1 ' paste �����y�[�W�̍ŏ��̃y�[�W
    'sttIn    = pTo.Slides.Count + 1  ' paste �����y�[�W�̍ŏ��̃y�[�W
    sttIn    = pTo.Slides.Count  ' paste �����y�[�W�̍ŏ��̃y�[�W  ' 2014/03/24
    'MsgBox("sttIn = " & sttIn)
    sttIns(k)=sttIn
    If Desgns(k) = 0 Then
      Call copySlide(pFr,pTo,Pagesc(k),C_CMM)
    Else
      Call copySlide_Fmt(pFr,pTo,Pagesc(k),C_CMM)
    End If
  Next k

  ' ���� text file �̍쐬
  Dim sItgFilePath   As String
  'Call notesFilePath(sItgFilePath,sCurrentFolder,"_intg.txt",0,1)
  Call notesFilePath(sItgFilePath,sCurrentFolder,".txt",0,GFLAG) ' 2014/04/06
  Call offsetInteg(Fnames(),Pagesc(),C_CMM,sttIns(),sItgFilePath)

  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

End Sub

Public Sub grayTOC(GFLAG As Integer)
  ' �K�w�ڎ��̎w���������ꂽ�ݒ�t�@�C��(text file)��ǂݍ���,
  ' �K�w�ڎ�(�D�F)��R��ׂ��y�[�W�ɒǉ����ĕʃt�@�C���Ƃ��ĕۑ�����.
  ' �ǉ�����ƃy�[�W���������̂�, �����␳���ĕۑ������ʃt�@�C���̐ݒ�t�@�C���Ƃ��ĕۑ�����.
  ' ��: hoge.ppt �ɊK�w�I�ڎ���ǉ�����. �ݒ�t�@�C���� hoge.txt
  ' hoge.txt �̒��g:
  ' ::|||::
  ' �{���̓��e
  ' *  1. �w�i              1
  ' ** 1.1. �͂��߂�        2
  ' ** 1.2. ���j            3
  ' *  2. �ړI              4
  ' hoge.txt �� format:
  ' ::|||::
  ' *   (\t)(�K�w1�̖ڎ�)(\t)(\t)(���������y�[�W��)�@��: ���̃y�[�W�̑O�ɒǉ�����. ���̃y�[�W�ȍ~�̃y�[�W���������.
  ' **  (\t)(�K�w2�̖ڎ�)(\t)(\t)(���������y�[�W��)
  ' *** (\t)(�K�w3�̖ڎ�)(\t)(\t)(���������y�[�W��)
  '
  ' output:
  ' �K�w�ڎ����ǉ����ꂽ ppt file: hoge_grayTOC.ppt
  ' �ǉ��K�w�ڎ��̃y�[�W���������␳, �ڎ����J�E���g���Ȃ��X���C�h�Ƃ��Ēǉ����� txt file: hoge_grayTOC.txt
  '

  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,GFLAG)

  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  ' iNc***(1 ���� hire, 1 ���� page ��) �̓񎟌��z��
  ' iNcIdx(hier,page) = ���̃y�[�W�̑O�ł�, ���� id �� Toc �������������܂�(���͊D�F)
  ' iNcTtl(hier,page) = ���̃y�[�W�̑O��, ���̃^�C�g���Ŗڎ��������܂�.
  ' iNcToc(hier,page) = ���̃y�[�W�̑O�ɏ����ڎ�����(���s�� split ����Ɣz��ɂȂ�)
  ' iNcOff(hier,page) = ���̃y�[�W�̑O�ɒu���ڎ��y�[�W�̐� (offset ���v�Z����Ƃ��Ɏg��)
  Dim iNcIdx()  As String
  Dim iNcTtl()  As String
  Dim iNcToc()  As String
  Dim iNcOff()  As Integer
  Call bufRead_TOC(hBufs.Item(GTOC_MARK),iNcIdx(),iNcTtl(),iNcToc(),iNcOff(),C_TBB)

  Dim sPptFileName As String
  sPptFileName=newFileName("_gray.pptm","do you want to save the file As ",GFLAG)
  ' �ʖ��ŏ㏑���ۑ�
  ' ��������͂��̃t�@�C���� Active �ƂȂ�̂Œ��ӂ���.
  ' [2018-02-06 Tue]
  Dim yvb As Integer
  If GFLAG = 1 Then
    yvb = MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo)
  Else
    yvb = vbYes
  End If
  'If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
  If yvb = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

  Call add_grayTOC(iNcIdx(),iNcTtl(),iNcToc(),iNcOff())

  ' page �̃Y�����v�Z����
  ' �V�����y�[�W�� = �Â��y�[�W�� + offset(�Â��y�[�W��)
  Dim offset() As Integer
  Call offset_grayTOC(offset(), iNcOff())

  Dim sGrayFilePath As String
  Dim sGrayFileName As String
  sGrayFileName=newFileName(".txt","do you want to save the file As ",GFLAG)
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
        If InStr(oShape.Name,C_TITLE1) Then     ' 2014/10/08 oShape.Name �� "Title" ���܂�ł���ύX
          oShape.TextFrame2.TextRange.Characters.Font.Size=Tsize
        End If
        If InStr(oShape.Name,C_PLCHOLDER) Then
          ' �K���ɂ������ł���. ����ł����̂�?
          oShape.TextFrame2.TextRange.Characters.Font.Size=Fsize
          'oShape.TextFrame.TextRange.Characters.Font.Color.RGB=Blk
        End If
      Next oShape
    End If
  Next i
End Sub

Public Sub removeGrayTOC_asNewName(GFLAG As Integer) 'for print (pdf) hand out
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
  spptFileName = newFileName("_delTOC.pptm","do you want to save the file as ",GFLAG)
  Dim yvb As Integer
  If GFLAG = 1 Then
    yvb = MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo)
  Else
    yvb = vbYes
  End If
  'If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
  If yvb = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

End Sub

' hBufs �� offset �ɂ��ƂÂ��ĕ␳���� hash string (nBufs) ��Ԃ��܂�.
Private Sub offsetHash(hBufs As Object, nBufs As Object, offset() As Integer, Optional pExist As Variant)

  Dim keys() As Variant
  keys = hBufs.Keys
  Dim key As Variant

  For Each key In keys
    Dim sBuf As String
    If key = GTOC_MARK Then
      sBuf=""  ' ���ł��悭�킩��Ȃ�. scope �������ĂȂ��̂� ?
      Call bufWrit_TOC(sBuf,hBufs.Item(key),C_TBB,offset(),pExist)
    ElseIf key = CPPT_MARK Then
      sBuf=""
      Call bufWrit_File(sBuf,hBufs.Item(key),C_TBB,offset(),pExist)
    ElseIf key = TOCS_MARK Or key = TOHI_MARK Then
      sBuf=""
      Call bufWrit_pList(sBuf,hBufs.Item(key),C_TBB,offset(),pExist)
    ElseIf key = NOWT_MARK Or key = NOHI_MARK Or key =DELS_MARK Or key = ALIV_MARK Or key = SKIP_MARK Then
      sBuf=""
      Call bufWrit_nList(sBuf,hBufs.Item(key),C_CMM,offset(),pExist)
    ElseIf key = FLIP_MARK Then
      sBuf=""
      Call bufWrit_exList(sBuf,hBufs.Item(key),C_TBB,offset(),pExist)
    ElseIf key = IFNT_MARK Or key = IFHI_MARK Then
      sBuf=""
      Call bufWrit_siz(sBuf,hBufs.Item(key))
    End If
    If nBufs.Exists(key) Then
      nBufs.Item(key) = nBufs.Item(key) & vbNewLine & sBuf
    Else
      nBufs.Add key,sBuf
    End If
  Next key

End Sub

' offset �␳���� text file �������o���܂�.
Private Sub offsetPrint(sFilePath As String, hBufs As Object, offset() As Integer, Optional pExist As Variant)
  '''Set nBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim nBufs As Object
  Set nBufs=CreateObject("Scripting.Dictionary")
  Call offsetHash(hBufs,nBufs,offset(),pExist)
  Call hashPrint(nBufs,sFilePath)
End Sub

' hash �œ���ꂽ String �����܂��� format �ŏ����o���܂�
Private Sub hashPrint(hash As Object, sFilePath As String)
  Dim keys() As Variant
  keys = hash.Keys
  Dim key As Variant

  Dim prnt As String
  For Each key In keys
    prnt = prnt & INFO_MARK & key & INFO_MARK & vbNewLine
    prnt = prnt & hash.Item(key)         & vbNewLine
    'MsgBox("key= " & key & vbTab & "value = " & hash.Item(key))
  Next key

  Dim iFileNum As Integer
  iFileNum = FreeFile()
  Open sFilePath For Output As #iFileNum
  Print #iFileNum, prnt
  Close #iFileNum
End Sub

' gray �ڎ���}���������Ƃɂ�� offset �l���v�Z���܂�.
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

' text file �̏��Ɋ�Â��� gray �ڎ��X���C�h��}�����܂�.
Private Sub add_grayTOC(iNcIdx() As String, iNcTtl() As String, iNcToc() As String, iNcOff() As Integer)
  Dim nhier As Integer
  nhier = UBound(iNcTtl,1) ' �K�w�̐�
  Dim npage As Integer
  npage = ActiveWindow.Presentation.Slides.Count ' �X���C�h�̖���

  If npage > UBound(iNcTtl,2) Then
    npage  = UBound(iNcTtl,2)
  End If

  ' �F�ݒ�
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
      ' �X���C�h�ǉ��ƊD�F���\�b�h�ŊK�w�ڎ������������o��
      If Not IsNull(title) And Not title = "" Then        ' �ڎ��^�C�g�������݂����
        If page <= npage Then ' �y�[�W�������
          Dim pageo As Integer
          pageo = page + offs
          'MsgBox("hier=" & hier & " page=" & page & " offs=" & offs & " pageo=" & pageo)
          ' (1) �D�F���\�b�h�ł̖ڎ��X���C�h���܂��ǉ�
          ActivePresentation.Slides.Add pageo, ppLayoutText ' ���̑O�ɃX���C�h�ǉ�
          ActivePresentation.Slides(pageo).Shapes(1).TextFrame.TextRange=title ' title ������
          ' �ڎ��������v���[�X�z���_�[
          Set oTxtRng = ActivePresentation.Slides(pageo).Shapes(2).TextFrame.TextRange
          Dim toc As Variant
          ' +2 �����y�[�W�ɕ����̍��ڂ̐��������邱�Ƃ�z�肷��ꍇ(���� hierarchy �Ńy�[�W��������)
          '''Set hsh=New Dictionary  ' +2
          Dim hsh As Object
          Set hsh=CreateObject("Scripting.Dictionary")
          Dim k As Variant                              ' +2
          'MsgBox("hier=" & hier & " page=" & page & " idx=" & iNcIdx(hier,page))
          For Each k In Split(iNcIdx(hier,page),vbNewLine)                ' +2
            If Not k = "" Then     ' ���̂���Ȓl������̂��悭�킩��Ȃ�  '+2
              hsh.Add CInt(k),1                                            '+2
            End If                                                         '+2
          Next k                                                           '+2
          Dim idx As Integer
          idx = 0
          Dim sname As String
          sname  = C_GRAYTOC & "_" & title & "_" & CStr(hier) & "_" & CStr(page)
          Dim sname2 As String
          sname2=""
          For Each toc In Split(iNcToc(hier,page),vbNewLine) '��s������
            With oTxtRng.Paragraphs(idx) '�ӏ�����1����
              .Text = toc & vbNewLine
              ' �������̃y�[�W�Ő������鍀�� = ����
              'If iNcIdx(hier,page) = idx Then ' +1 �z�肵�Ȃ��ꍇ '+1
              If hsh.Item(idx) = 1 Then             ' �z�肷��ꍇ   '+2
                .Font.Color.RGB=Blk
                sname2 = sname2 & "_" & toc
              Else
                .Font.Color.RGB=Gry   ' �֌W�Ȃ����� = �D�F��
              End If
            End With
            idx = idx + 1
          Next toc
          hsh.RemoveAll ' ���e���폜  '+2
          ActivePresentation.Slides(pageo).Name = sname & sname2
          ' (2) id �� 0 �ł����, ���̑O�ɕ��ʂ̖ڎ���ǉ�
          'If InStr(iNcIdx(hier,page),"0") Then  ' ���ꂾ�� id=10 �̎����Y�����Ă��܂�!! 2014/05/10
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

' offset �l���v�Z���܂�.
' offset(page) �������ꍇ, page = page + offset(UBound(offset)) ���g���܂�
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

' offset �l�y�� exist �l(optional. �l������y�[�W���� 0 or 1 �ŋ��. exist(10)=1: 10 �y�[�W�ڂ͍l���ɓ���邱��)
' ��p����, �K�w�ڎ�(for �D�F�ڎ�)��V�������ď����o���܂�.
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
  Dim fst As Integer  ' �ŏ��̏����o���̃t���O(�����������s��)
  fst = 0
  For j = 0 To m
    If Left$(aBuf(j),1) = "*" Then
      Dim bf() As String
      bf = Split(aBuf(j),delm)    ' **(\t)�ڎ�(\t)(\t)�y�[�W��(���l)
      Dim page As Integer
      page = CInt(bf(UBound(bf))) ' �Ō�̃J�������y�[�W��(���l)
      'If pExist Is Nothing Then   ' pExist ����`����ĂȂ��Ƃ�
      'If IsMissing(pExist) Then   ' �����ʉ߂��� Else �ȉ��ŃG���[�ƂȂ邱�Ƃ�����
      'If IsMissing(pExist) Or pExist Is Nothing Then ' ����ő��v��? 2014/03/27
      ' pExist ���z��Ƃ��đ��݂��Ă���Ƃ�,
      ' If pExist Is Nothing Then ' ����̓G���[�ƂȂ� Why?
      If IsMissing(pExist) Or IsEmpty(pExist) Or IsNull(pExist) Then ' 2014/04/06 ad hoc
        page = page + oFFst(page,offset()) ' page �������炷
        bf(UBound(bf)) = CStr(page)
        prnt = prnt & join(bf,delm) & vbNewLine
        fst= fst+1
      Else
        'If Not page > UBound(pExist) AndAlso pExist(page) = 1 Then    ' 2014/04/21 And -> AndAlso
        ' 2014/04/21 �Z���]�����ł��Ȃ�?
        If Not page > UBound(pExist) Then
          If pExist(page) = 1 Then
            page = page + oFFst(page,offset()) ' page �������炷
            ' �y�[�W�̓r������n�܂��Ă���Ƃ���, �K�w�����܂��L�q�ł��Ȃ�
            ' �ŏ��̏��������̍ۂɊK�w�����ǂ��ď����Ă���.
            ' 2014/04/25
            If fst = 0 Then
              Dim star As Integer
              star =Len(bf(0))
              If star > 1 Then      ' �K�w�� '*' �łȂ��Ƃ�
                Dim spx() As String ' �K�w���ɕ���������Ă�������������.
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
                For s = 1 To star-1 ' �K�w�����ǂ���
                  If Not IsNull(spx(s)) Then
                    If Not spx(s) = "" Then
                      prnt = prnt & spx(s) & vbNewLine ' �ŏ��̊K�w�ڎ�������
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

' �K�w�ڎ��̍\����ǂݍ��݂܂�.
Private Sub bufRead_TOC(sBuf As Variant, iNcIdx() As String, iNcTtl() As String, iNcToc() As String, iNcOff() As Integer,delm As String)
  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine) ' ��s������

  Dim m As Integer
  m = UBound(aBuf) ' m = �s�� -1 (0-based)

  ' �K�w�̐��𐔂���(* �̐��𐔂���)
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
  For h = 1 To hier ' �K�w�̐�. �K�w��������Ă���
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
        page =CInt(bf(UBound(bf)))  ' �Ō�̗v�f = �y�[�W��
        star =Len(bf(0))            ' 2014/03/11 tab ��؂�ɕύX
        toc  =bf(1)
        'cf   =Split(bf(0))          ' ������ space split
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
            ' +2 �����y�[�W�ɕ����̍��ڂ̐��������邱�Ƃ�z�肷��ꍇ
            If IsNull(iNcIdx(h,p)) Or iNcIdx(h,p) ="" Then  '+2
              iNcOff(h,p)= iNcOff(h,p) + 1                  '+2
              If id = 0 Then iNcOff(h,p)= iNcOff(h,p) + 1   '+2
            End If                                          '+2
            iNcIdx(h,p)=iNcIdx(h,p) & id & vbNewLine        '+2

            ' +1 �z�肵�Ȃ��ꍇ
            ' �����y�[�W�ɓ��� hierarchy �̐����𕡐����Ȃ��̂ł���΂��������g�������ǂ�
            ' +2, +1 ����ւ�
            'iNcOff(h,p)= iNcOff(h,p) + 1                   '+1
            'If id = 0 Then iNcOff(h,p)= iNcOff(h,p) + 1    '+1
            'iNcIdx(h,p)= id                                '+1

            iNcTtl(h,p)=title
            iNcToc(h,p)=join(tocs,vbNewLine) ' �񎟌��z��ʓ|�Ȃ̂ŉ��s��

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

' text file ��ǂݍ���� hash �Ɋi�[���܂�
' String �Ƃ��ēǂݍ��ނ���.
' nock �� 1 �̎���, �t�@�C���̑��݂��m�F���Ȃ�(������΂��̂܂܏I���)
Private Sub fileRead2(sNotesFilePath As String, hBufs As Object, Optional nock As Integer)

  If nock = 1 And Dir(sNotesFilePath) = "" Then
    Exit Sub
  End If

  ' �t�@�C����ǂݏo��
  Dim iNotesFileNum As Integer
  iNotesFileNum = FreeFile()
  Open sNotesFilePath For Input As iNotesFileNum

  Dim MARK As String
  Dim sBuf As String
  Do Until EOF(iNotesFileNum)
    Line Input #iNotesFileNum,sBuf  ' sBuf �̒��Ɉ�s������Ă���
    If sBuf = "" Then GoTo SKIP                 ' ��s�͓ǂݔ�΂�
    If left$(sBuf, 1) = COMT_MARK Then GoTo SKIP  ' �R�����g�s���������΂�
    If Left$(sBuf,2) = INFO_MARK And Right$(sBuf,2) = INFO_MARK And Len(sBuf)= Len(INFO_MARK)*2+3 Then '���ڈ�
      MARK = Mid$(sBuf,3,3)
    Else
      If Not IsNull(MARK) And Not MARK = "" Then
        If hBufs.Exists(MARK) Then
          hBufs.Item(MARK) = hBufs.Item(MARK) & vbNewLine & sBuf
        Else
          hBufs.Add MARK,sBuf
        End If
      End If
    End If
SKIP:
  Loop
  Close iNotesFileNum

End Sub

' text file �����擾���܂�.
' default �ł�, hoge.pptx => hoge.txt
' ck = 1   �̏ꍇ, �t�@�C����������Α��v���O�������I��.
' mflg = 1 �̏ꍇ, MsgBox �� default �� text file ���g�����ǂ������m�F.
Private Sub notesFilePath (sNotesFilePath As String, sCurrentFolder As String, suffix As String, ck As Integer, mflg As Integer, Optional Fname As String)

  Dim delm As String
  delm = "\"   ' default (windows)

  Dim Op As Variant
  Op = Application.OperatingSystem
  If Op Like "Macintosh*" Then delm = ":"

  Dim sNotesFileName As String   ' text file ��
  ' �ݒ� text file
  sCurrentFolder = ActivePresentation.Path & delm
  ' mflg = 1 �m�F MsgBox ��\������
  sNotesFileName = newFileName(suffix,"do you want to use the file ",mflg, Fname)

  sNotesFileName = Replace(sNotesFileName,"\",delm)

  'MsgBox("Fname=" & Fname)
  'MsgBox("sNotesFileName=" & sNotesFileName)
  'MsgBox("sCurrentFolder=" & sCurrentFolder)

  If InStr(sNotesFileName, delm) = 0 Then
    sNotesFilePath = sCurrentFolder & sNotesFileName
  'ElseIf Left$(sNotesFileName,2) = "." & delm Then  ' 2014/04/11 �폜
  '  sNotesFilePath = sCurrentFolder & Mid(Fname,2)
  ElseIf Left(sNoteFileName,1) = "." Then  ' 2014/04/11 ����ς�悭�킩��Ȃ�.
    sNotesFilePath = sCurrentFolder & sNoteFileName
  Else
    sNotesFilePath = sNotesFileName  ' ��΃p�X�ŏ����Ă���ꍇ
    If Op Like "Macintosh*" Then     ' Mac �̏ꍇ(�����������@�Ȃ��񂾂�[��?)
      If InStr(sNotesFilePath, "Macintosh")=0 Then
        sNotesFilePath = "Macintosh HD" & sNotesFilePath
      End If
    End If
  End If
  ' is it there? quit if not
  'sNotesFilePath = sCurrentFolder & sNotesFileName
  'MsgBox("sNotesFilePath=" & sNotesFilePath)

  If ck = 1 Then  ' �t�@�C���̑��݂��m�F����ꍇ ck = 1
    If Len(Dir$(sNotesFilePath)) = 0 Then
      MsgBox (sNotesFilePath & " is missing")
      End
    End If
  End If

End Sub


' �������� text file �ɂ��� offset �␳��,
' �����e�L�X�g�t�@�C�����쐬���܂�.
Private Sub offsetInteg (Fname() As String, Pagesc() As String, delm As String, sttIns() As Integer, sFilePath As String)
  '''Set nBufs = New Dictionary ' hash
  Dim nBufs As Object
  Set nBufs=CreateObject("Scripting.Dictionary")
  Dim m As Integer
  m = UBound(Fname)
  For k = 0 To m
    Call offset_Intg(nBufs,Fname(k),Pagesc(k),delm,sttIns(k))
  Next k
  Call hashPrint(nBufs,sFilePath)
End Sub

' ��荞�ރy�[�W�̃��X�g�ƌ��݂̃y�[�W������,
' offset() �̒l�� pExit() ���v�Z���܂�.
Private Sub offset_Intg(nBufs As Object, Fname As String, Pagesc As String, delm As String, sttIn As Integer)

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

  ' pExist(�֌W����y�[�W��) = 1
  ' pExist(�Ȃ��y�[�W��)     = 0
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
  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs,1)  ' �Ō�̈��� =1 �t�@�C�������Ă����̂܂ܐi��.
  Dim Prng As String
  Call joinInt(Prng,pExist(),",")
  Call offsetHash(hBufs,nBufs,offset(),pExist)

End Sub

' collectPpt �p�̏��������o���܂�.
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
    If sForm <> "" Then  ' 2014/04/29 �y�[�W���������ꍇ�̒u��������͏����Ȃ�
      prnt = prnt & Desgns(k) & delm & Fnames(k) & delm & sForm & vbNewLine
    End If
  Next k

End Sub

' collectPpt �p�̏���ǂݍ��݂܂�
Private Sub bufRead_File(sBuf As Variant, Desgns() As Integer, Fnames() As String, Pagesc() As String, delm As String)
  Dim Op As Variant
  Op = Application.OperatingSystem

  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine) ' ��s������
  'aBuf = Split(sBuf,vbCrLf)

  Dim fnum     As Integer
  Dim Pdummy() As String
  fnum = 0
  Dim tBuf As Variant
  For Each tBuf In aBuf
    ReDim Preserve Desgns(fnum)
    ReDim Preserve Fnames(fnum)
    ReDim Preserve Pdummy(fnum)
    Dim buf() As String
    buf = Split(tBuf,delm) ' split by vbTab(�Œ�)
    ' 1. �X�^�C�����R�s�[���邩�ǂ��� (0 or 1)
    Desgns(fnum)=CInt(buf(0))
    ' 2. ppt file name
    Dim Fname As String
    ' 2014/06/27 �֐��ɂ���
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
'      ' ����Fname ���t�@�C���������̏ꍇ��, �J�����g�f�B���N�g���ɂ���Ƃ݂Ȃ���
'      ' ��΃p�X�ɂ���.
'       '���邢�� "." ����n�܂�ꍇ�ɂ͑��΃p�X�ŏ����Ă���ƍl����
'      ' ������f�B���N�g������ǋL���Đ�΃p�X�ɂ���
'      ' ����ȊO�̏ꍇ��, ��΃p�X�ŏ����Ă���Ƃ��񂪂���.
'      If InStr(Fname, "\") = 0 Then
'        Fname = ActivePresentation.Path & "\" & Fname
'      'ElseIf Left$(Fname,2) = ".\" Then   ' 2014/04/11 �폜
'      '  Fname = ActivePresentation.Path & "\" & Mid(Fname,2)
'      ElseIf Left(Fname,1) = "." Then ' 2014/04/11 �悭�킩��Ȃ�.
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
    Set Pptf = Presentations.Open(FileName:=Fnames(k),ReadOnly:=msoFalse) ' �g���X���C�h���Ƃ肠�����J��
    Dim defList() As Integer
    Call mkDefList(defList(),Pptf)
    Dim npage As Integer
    npage = Pptf.Slides.Count
    Pptf.Close
    ' ad hoc 2014/04/11
    'Pdummy(k) = "2-" �̂悤��, "-" �ŏI����Ă���ꍇ�ɂ�, npage ��ǉ�
    'Pdummy(k) = "2-18"
    If Right(Pdummy(k),1) = "-" Then
      Pdummy(k) = Pdummy(k) & CStr(npage)
    End If
    Call sForm2sList(Pdummy(k),Pagesc(k),deflist(),C_CMM)
  Next k
End Sub

' slide ���R�s�[���܂�.
Private Sub copySlide(pFr As Presentation, pTo As Presentation, Pagesc As String, delm As String)
  Dim Pages() As Integer
  Call sList2iList(Pagesc,Pages(),C_CMM)
  ' copy slide (as usual)
  pFr.Slides.Range(Pages).Copy
  pTo.Slides.Paste
  pFr.Close
End Sub

' slide �̃X�^�C�������� slide �̏�Ԃɕۂ����܂܃R�s�[���܂�
Private Sub copySlide_Fmt(pFr As Presentation, pTo As Presentation, Pagesc As String, delm As String)

  Dim Pages() As Integer
  Call sList2iList(Pagesc,Pages(),C_CMM)

  ' keep formatting copy
  ' http://img2.tapuz.co.il/forums/1_88724584.txt
  Dim q  As Integer
  For q = 0 To UBound(Pages) ' 1 page ���R�s�[����
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


' ������ countbox ���쐬���܂�.
' �����y�[�W "::---::"
' �y�� gray �ڎ��y�[�W��
' �������܂�.
Public Sub pageCountsBox2(GFLAG As Integer)

  ' text file ��ǂݍ���� page �ԍ���t���Ȃ����̂��w��
  ' page �ԍ���t���Ȃ�����: "::---::" �ŏ����ꂽ�y�[�W
  ' �X���C�h���� grayTOC �Ŏn�܂����

  ' �F�ݒ�
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

  ' �ꏊ�ݒ�
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
  Dim npage0  As Integer  ' �X���C�h����
  Dim npage   As Integer  ' skip ���������X���C�h����
  Call skipSlides(iSkip(), npage0, npage, GFLAG) ' [2018-02-06 Tue] add GFLAG

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
        ' �g��
        If (j Mod 10 = 0) Then
          oSp.Line.ForeColor.RGB=Red
        Else
          oSp.Line.ForeColor.RGB=Blk
        End If
        ' �h��Ԃ�
        If (j < page) Then
          oSp.Fill.ForeColor.RGB=Gry
        ElseIf (j = page) Then
          oSp.Fill.ForeColor.RGB=Col
        Else
          oSp.Fill.ForeColor.RGB=Wht
        End If
        ShapeList(j) = Shp.Name
      Next
      ' �O���[�v�����Ė��O������
      ActivePresentation.Slides(i).Shapes.Range(ShapeList()).Group.Name = C_COUNTBOX
    End If
  Next
  ActivePresentation.Slides(1).Select
  ' ���������
End Sub

Public Sub removeCountsBox()
  Dim SlideObj   As Slide
  Dim ShapeObj   As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoGroup Then            ' �O���[�v
        If ShapeObj.Name = C_COUNTBOX Then      ' ���O�� "countbox"
          ShapeObj.Delete
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub


' �y�[�W�ԍ���t���܂�
' �����y�[�W "::---::"
' �y�� gray �ڎ��y�[�W��
' �������܂�.
Public Sub pageNum2(GFLAG As Integer)

  ' text file ��ǂݍ���� page �ԍ���t���Ȃ����̂��w��
  ' page �ԍ���t���Ȃ�����: "::---::" �ŏ����ꂽ�y�[�W
  ' �X���C�h���� grayTOC �Ŏn�܂����

  ' �ꏊ�w��
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
  Dim npage0  As Integer  ' �X���C�h����
  Dim npage   As Integer  ' skip ���������X���C�h����
  Call skipSlides(iSkip(), npage0, npage, GFLAG) ' iSkip(page)=1 �̂Ƃ�, page �Ԗڂ̃X���C�h���X�L�b�v����

  Dim page As Integer
  page = 0
  Dim i As Integer
  Dim Shp
  For i = 1 To npage0
    If Not iSkip(i) = 1 Then
      page = page + 1
      ActivePresentation.Slides(i).Select
      ' i ���ڂ̃X���C�h
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
      Shp.Name = C_PAGENUM  ' textbox �ɖ��O�����Ă���
    End If
  Next i
End Sub

Public Sub removePageNum()
  Dim SlideObj As Slide
  Dim ShapeObj As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides          ' �e�X���C�h
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1   ' �e�V�F�C�v
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoTextBox Then              ' TextBox �ł���
        If ShapeObj.Name = C_PAGENUM Then             ' ���O�� counttxt �ł���.
          ShapeObj.Delete                             ' �I�u�W�F�N�g����������.
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub

' ��������y�[�W���v�Z���܂�
Private Sub skipSlides(iSkip() As Integer, ByRef npage0 As Integer, ByRef npage As Integer, mflg As Integer) ' [2018-02-06 Tue] add gFlag

  ActivePresentation.Slides(1).Select     ' �ꖇ�ڂ̃X���C�h��I��
  npage0 = ActivePresentation.Slides.Count ' �X���C�h�̑���(�S��)

  ReDim iSkip(npage0)
  Dim isk As Integer
  isk = 0

  ' [2018-02-06 Tue]
  Dim yvb As Integer
  If mflg = 1 Then
    yvb = MsgBox("do you want to use text file for skip ?",vbYesNo)
  Else
    yvb = vbYes
  End If

  ' "::---::" �t���O���e�L�X�g�t�@�C������Ƃ��Ă��邩�ǂ���
  'If MsgBox("do you want to use text file for skip ?",vbYesNo) = vbYes Then
  If yvb = vbYes Then
    ' text file(path) �̎擾
    Dim sNotesFilePath As String
    Dim sCurrentFolder As String
    Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",0,mflg) ' [2018-02-06 Tue]
    '''Set hBufs = New Dictionary ' �ǂݍ��񂾕�����
    Dim hBufs As Object
    Set hBufs=CreateObject("Scripting.Dictionary")
    Call fileRead2(sNotesFilePath,hBufs)
    Dim nSids()  As Integer
    Call bufRead_nList(hBufs.Item(SKIP_MARK),nSids(),C_CMM) ' vbTab -> ","
    Dim k As Integer
    If Not Sgn(nSids) = 0 Then  ' 2014/03/27
      For k = 0 To UBound(nSids)
        iSkip(nSids(k))=1
        isk = isk + 1
      Next k
    End If
  End If

  ' �y�[�W���Ƃ��ăJ�E���g���Ȃ��ڎ��̖����𐔂���
  Dim i As Integer
  For i = 1 To npage0
    If InStr(ActivePresentation.Slides(i).Name, C_GRAYTOC) Then
      iSkip(i)=1
      isk = isk + 1
    End If
  Next i
  npage = npage0 - isk
End Sub

' �e�X�̃X���C�h���ɖڎ����쐬��,
' ���݂ǂ��ɂ��邩�𖾎����܂�.
Public Sub pageListContents2(GFLAG As Integer)

  ' ppt �̖ڎ��������ŏ����o��.
  ' �e�X�̃y�[�W(�E��)��, ������Ă�ڎ��̈ʒu�������F��,
  ' ���̑��̖ڎ��̈ʒu���D�F�ŏ����o��
  ' �ڎ���, powerpoint �t�@�C�����Ɠ����̃e�L�X�g�t�@�C���ɏ����Ă���.
  ' ��: hoge.ppt �̖ڎ��� hoge.txt �ɏ����Ă���.
  ' hoge.txt �̒��g:
  ' ::###::
  ' 1. �w�i              1-3
  '    1.1. �͂��߂�     2
  '    1.2. ���j         3
  ' 2. �ړI              4-6
  '...
  ' ::$$$::
  ' 1,2,3,4
  ' hoge.txt �̒��g�I��
  '
  ' hoge.txt �� format:
  '::###::
  ' (�����o�����e)\t\t\t(�y�[�W��),(�y�[�W��)-(�y�[�W��)
  '::$$$::
  '(�y�[�W��),(�y�[�W��)-(�y�[�W��)...     <- �ڎ��������o���Ȃ��y�[�W��
  ' ...
  '
  '
  ' 2013/05/16 option (�����ȉ��̏�񂪂���΂�����g��. ������� MsgBox �ŕ����Ă���)
  ' 2014/03/01 format �̕ύX
  ' ::!!!::
  '(�ڎ���u���ꏊ��x���W)\t(�ڎ�textbox�̕�)\t(�t�H���g�T�C�Y)

  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,GFLAG)

  ' file reading
  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  ' �ڎ��Ƃ��ď����������e�ƃy�[�W��
  Dim Conts()  As String
  Dim Pagesc() As String
  If Not IsNull(hBufs.Item(TOCS_MARK)) And Not hBufs.Item(TOCS_MARK) = "" Then
    Call bufRead_pList(hBufs.Item(TOCS_MARK),Conts(),Pagesc(),C_TBB)
  ElseIf Not IsNull(hBufs.Item(GTOC_MARK)) And Not hBufs.Item(GTOC_MARK) = "" Then
    Call bufRead_TOC2pList(hBufs.Item(GTOC_MARK),Conts(),Pagesc(),C_TBB)
  End If

 ' �ڎ��������o���Ȃ��X���C�h�ԍ�(1-based)
  Dim nSids()  As Integer
  Call bufRead_nList(hBufs.Item(NOWT_MARK),nSids(),C_CMM) ' vbTab -> ","

  ' ���������ꏊ
  Dim xPos0  As Integer
  Dim wid0   As Integer
  Dim Fsize0 As Integer
  Dim sizFlg As Integer
  ' �����l����Ă���
  xPos0  = C_XPOSR
  wid0   = C_WIDER
  Fsize0 = C_FSIZE
  sizFlg = 0
  If Not IsNull(hBufs.Item(IFNT_MARK)) And Not hBufs.Item(IFNT_MARK) = "" Then
    Call bufRead_siz(hBufs.Item(IFNT_MARK),xPos0,wid0,Fsize0)
    sizFlg = 1
  End If
  If GFLAG = 0 Then  ' [2018-02-06 Tue]
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
  hei   = 20    ' �֌W�Ȃ�
  'Fsize = 10
  xPos  = C_XPOSR
  yPos  = C_YPOSR
  wid   = C_WIDER
  Fsize = C_FSIZE

  ' �����o���ꏊ, ��, �t�H���g�T�C�Y�̕ύX
  Call siz_MsgBox(sizFlg,xPos,wid,Fsize,xPos0,wid0,Fsize0)

  ' �X���C�h�֏����o��
  Call write_TOC(xPos,yPos,wid,hei,Fsize,nSids(),Conts(),Pagesc(),C_CONTENT,1, GFLAG)

End Sub

' ���ۂɖڎ��������o���܂�
Private Sub write_TOC(xPos As Integer,yPos As Integer,wid As Integer,hei As Integer,Fsize As Integer,nSids() As Integer,Conts() As String,Pagesc() As String,sName As String, aFlg As Integer, mflg As Integer)

  Dim iSkip() As Integer
  Dim npage0  As Integer  ' �X���C�h����
  Dim npage   As Integer  ' skip ���������X���C�h����
  Call skipSlides(iSkip(),npage0,npage,mflg)

  Dim fSld() As Integer ' ���������X���C�h�ԍ� i fSld(i)=1
  Call checkw_slde(nSids(),fSld(),npage0)

  Dim cnum As Integer   ' �ڎ����ڐ�
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
      ' i ���ڂ̃X���C�h
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
          cont = Conts(j) & ncnt & vbNewLine ' �Ō�ȊO�͉��s������
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
      Shp.ZOrder (msoSendToBack)  ' �ŉ��w�ɒu��
      Shp.Name = sName
    End If
  Next

End Sub

Private Sub bufRead_TOC2pList(sBuf As Variant,Conts() As String, Pagesc() As String, delm As String)
  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine) ' ��s������

  Dim m As Integer
  'm = UBound(aBuf)-1 ' m = �s�� -2 (0-based) ' �ŏ��̍s(* ���Ȃ�)�͍l���Ȃ�.
  m = UBound(aBuf)  ' �ŏ��̍s�������Ƃ�������̂� 2013/03/29

  Dim hier   As Integer
  Dim tBuf() As String
  Dim j      As Integer
  hier = 1
  For j = 0 To m
    tBuf = Split(aBuf(j))
    If hier< Len(tBuf(0)) Then hier=Len(tBuf(0))
  Next j

  '''Set hbef = New Dictionary
  Dim hbef As Object
  Set hbef=CreateObject("Scripting.Dictionary")
  Dim stt() As Integer
  Dim edd() As Integer
  Dim sForm() As String
  For j = 0 To m
    Dim buff As String
    'buff = aBuf(j+1)
    buff = aBuf(j)    ' 2013/03/29
    If Left$(buff,1) = "*" Then
      ' ���I�ɕύX 2013/03/29
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
      page      =CInt(bf(UBound(bf)))  ' �Ō�̗v�f = �y�[�W��
      star      =Len(bf(0))            ' 2014/03/11 tab ��؂�ɕύX
      toc       =bf(1)
      'cf        =Split(bf(0))          ' ������ space split
      'star      =Len(cf(0))            ' �K�w�̐�
      'toc       =Mid(bf(0),star+2)     ' �ڎ�

      Conts(j)  =toc
      stt(j)    =page
      Dim st As Integer
      For st = star To hier
        If hbef.Exists(st) Then
          If Not stt(hbef.Item(st)) = page Then
            edd(hbef.Item(st))=page-1
          Else
            edd(hbef.Item(st))=C_EXC
          End If
          hbef.Item(st) = j
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
      If Not stt(hbef.Item(st)) = page Then
        edd(hbef.Item(st))=page-1
      Else
        edd(hbef.Item(st))=C_EXC
      End If
      hbef.Item(st) = j
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

' ���������X���C�h�ԍ� k: fSld(k) = 1
' nSids(i)=1: i �Ԗڂ͏��������Ȃ�
Private Sub checkw_slde(nSids() As Integer, fSld() As Integer, npage As Integer)
  ' �����o���Ȃ��X���C�h nSids() = ��: (1,2,5)
  ' i ���ڂ̃X���C�h�ɖڎ��������o��: fSld(i)=1
  ' �����o���Ȃ�: ��: fSld(1,2,5)=0
  ReDim Preserve fSld(1 To npage)
  Dim i As Integer
  For i = 1 To npage
    fSld(i) = 1       ' ��{�� 1
  Next i
  If Not Sgn(nSids) = 0 Then
    Dim j As Variant
    For Each j In nSids
      fSld(j) = 0       ' flag �������Ă��� = 0
    Next j
  End If
End Sub

' ���������X���C�h�ԍ� k: fSld(k) = 1
' nSids(i)=1: i �Ԗڂ���������
Private Sub checkd_slde(nSids() As Integer, fSld() As Integer, npage As Integer)
  ReDim Preserve fSld(1 To npage)
  Dim i As Integer
  For i = 1 To npage
    fSld(i) = 0      ' ��{�� 0
  Next i
  If Not Sgn(nSids) = 0 Then
    Dim j As Variant
    For Each j In nSids
      fSld(j) = 1      ' flag �������Ă��� = 1
    Next j
  End If
End Sub

' (�ڎ����e)(\t)(�y�[�W���̃��X�g)
' ����������
' (�y�[�W���̃��X�g)�ɂ�, offset()�y��, pExist() �ō��l���Ă���y�[�W�݂̂Ōv�Z���Ȃ���������
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
    If sForm <> "" Then '2014/03/24 �y�[�W���������ڎ��͏����Ȃ����Ƃɂ���.
      prnt = prnt & Conts(k) & delm & sForm & vbNewLine
    End If
  Next k

End Sub

'(�ڎ����e)(\t)(�y�[�W���̃��X�g) ��ǂݍ���
' (�y�[�W���̃��X�g)�� sList �`�� = "1,2,4,5,6,7,8" �̂悤�Ȋ����Ŏ����Ă���.
Private Sub bufRead_pList(sBuf As Variant, Conts() As String, Pagesc() As String, delm As String)

  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine)

  ' * �͎g��Ȃ�(dummy)
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

' (�y�[�W���̃��X�g)����������
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

' (�y�[�W���̃��X�g)��ǂݍ���
Private Sub bufRead_nList(sBuf As Variant, nSids() As Integer,delm As String)
  Dim aBuf() As String
  aBuf = Split(sBuf,vbNewLine)

  Dim sForm As String
  sForm = join(aBuf,delm)

  ' * �͎g��Ȃ�
  Dim defList(0) As Integer
  defList(0)=C_EXC

  Call sForm2iList(sForm,nSids(),delm,defList)

End Sub

' �������ވʒu, font ������������
Private Sub bufWrit_siz(prnt As String, sBuf As Variant)
  prnt = prnt + sBuf
End Sub

Private Sub bufRead_siz(sBuf As Variant, xPos0 As Integer, wid0 As Integer, Fsize0 As Integer)
  ' ��: format �̕ύX
  ' ::!!!::
  ' xPos0 \t wid0 \t Fsize0   <- sBuf

  Dim aBuf() As String
  aBuf = Split(sBuf,C_TBB)
  xPos0 =CInt(aBuf(0))
  wid0  =CInt(aBuf(1))
  Fsize0=CInt(aBuf(2))
End Sub

' font size ���̕ύX�� MsgBox ��
Private Sub siz_MsgBox(flg As Integer, x As Integer ,w As Integer ,f As Integer, x0 As Integer,w0 As Integer, f0 As Integer)
  If flg = 1 Then
    x=x0
    w=w0
    f=f0
  Else
    If MsgBox("x position= " & x & " textbox width =" & w & " font size = " & f & ": are they OK ?", vbYesNo) = vbNo Then
      x = InputBox("input x position of the textbox", "x �̒l��������=�������珑���o��", x)
      w = InputBox("input textbox width", "textbox width", w)
      f = InputBox("input font size", "font size", f)
    End If
  End If
End Sub

Public Sub removeListContents()
  Dim SlideObj As Slide
  Dim ShapeObj As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides          ' �e�X���C�h
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1   ' �e�V�F�C�v
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoTextBox Then           ' TextBox �ł���
        If ShapeObj.Name = C_CONTENT Then            ' ���O�� tcontents �ł���.
          ShapeObj.Delete                          ' �I�u�W�F�N�g����������.
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub


' �e�X�̃X���C�h���ɖڎ����쐬��,
' ���݂ǂ��ɂ��邩�𖾎����܂�.
Public Sub pageListHierarchy2(GFLAG As Integer)

  ' 2013/05/16
  ' ppt �̊K�w�\������������
  ' code from  pageListContents() �قƂ�ǂ��҂؂ō쐬(���Ƃ�����)
  ' 2014/04/09 �d�l�̕ύX
  ' ���܂܂�
  ' pageListHierarchy(^^^) -> pageListContents(###) -> grayTOC(|||) �̗D�揇��
  ' ������ȉ��ɕύX
  '  pageListHierarchy(^^^) -> grayTOC(|||) -> pageListContents(###) �̗D�揇��
  ' &&& ��ڈ�ɏ����o���Ȃ��y�[�W���w�肷��.
  ' ??? ��ڈ��, ���������ꏊ(x���W), textbox �̕�, font size ���w�肷��(option)

  '
  ' �e�X�̃y�[�W(����)��, ������Ă�ڎ��̈ʒu(�K�w)����������.
  ' �ڎ���, powerpoint �t�@�C�����Ɠ����̃e�L�X�g�t�@�C���ɏ����Ă���.
  ' ��: hoge.ppt �̖ڎ��� hoge.txt �ɏ����Ă���.
  ' �����Ȃ����, pageListContents() �̏������̂܂܎g��.
  ' hoge.txt �̒��g:
  ' ::^^^::
  ' 1. �w�i              1-3
  '    1.1. �͂��߂�     2
  '    1.2. ���j         3
  ' 2. �ړI              4-6
  ' ::&&&::
  ' 1-4,8
  ' hoge.txt �̒��g�I��
  '
  ' hoge.txt �� format:
  '::^^^::
  ' (�����o�����e)\t\t\t(�y�[�W��),(�y�[�W��)-(�y�[�W��)
  '::&&&::
  '(�y�[�W��),(�y�[�W��)-(�y�[�W��)...     <- �ڎ��������o���Ȃ��y�[�W��
  '
  '2013/05/16 option (�����ȉ��̏�񂪂���΂�����g��. ������� MsgBox �ŕ����Ă���)
  '2014/03/01 format �ύX
  '::???::
  '(�ڎ���u���ꏊ��x���W)\t(�ڎ�textbox�̕�)\t(�t�H���g�T�C�Y)
  '��:
  '::???::
  '10   130    8
  '
  ' ^^^ �������ꍇ�ɂ�, gray �ڎ�(|||)��p����.
  ' hoge.txt �̒��g:
  ' ::|||::
  ' *     1. �w�i            1-10
  ' **    1.1. �͂��߂�      1-3
  ' ***   1.1.1. �w�i(1)     1,2

  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,GFLAG)

  ' file reading
  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  ' �ڎ��Ƃ��ď����������e�ƃy�[�W��
  Dim Conts()  As String
  Dim Pagesc() As String
  ' 2041/04/09 �����̕ύX
  If Not IsNull(hBufs.Item(TOHI_MARK)) And Not hBufs.Item(TOHI_MARK)="" Then
    Call bufRead_pList(hBufs.Item(TOHI_MARK),Conts(),Pagesc(),C_TBB)
  ElseIf Not IsNull(hBufs.Item(GTOC_MARK)) And Not hBufs.Item(GTOC_MARK) = "" Then
    Call bufRead_TOC2pList(hBufs.Item(GTOC_MARK),Conts(),Pagesc(),C_TBB)
  ElseIf Not IsNull(hBufs.Item(TOCS_MARK)) And Not hBufs.Item(TOCS_MARK) = "" Then 
    Call bufRead_pList(hBufs.Item(TOCS_MARK),Conts(),Pagesc(),C_TBB)
  End If

   ' �ڎ��������o���Ȃ��X���C�h�ԍ�(1-based)
  Dim nSids()  As Integer
  Call bufRead_nList(hBufs.Item(NOHI_MARK),nSids(),C_CMM) ' vbTab -> ","

  ' ���������ꏊ
  Dim xPos0  As Integer
  Dim wid0   As Integer
  Dim Fsize0 As Integer
  Dim sizFlg As Integer
  xPos0  = C_XPOSL
  wid0   = C_WIDEL
  Fsize0 = C_FSIZE

  sizFlg = 0
  If Not IsNull(hBufs.Item(IFHI_MARK)) And Not hBufs.Item(IFHI_MARK) = "" Then
    Call bufRead_siz(hBufs.Item(IFHI_MARK),xPos0,wid0,Fsize0)
    sizFlg = 1
  End If
  'MsgBox("xPos0=" & xPos0 & vbNewLine & "wid0="  & wid0 & vbNewLine & "Fsize0=" & Fsize0)
  If GFLAG = 0 Then  ' [2018-02-06 Tue]
    sizFlg = 1
  End If

  Dim xPos  As Integer
  Dim yPos  As Integer
  Dim wid   As Integer
  Dim hei   As Integer
  Dim Fsize As Integer

  'xPos  = 590
  'xPos = 5  ' �K�w�̂Ƃ��͍���
  'yPos  = 5
  'yPos = 1
  'wid   = 125
  'wid   = 200
  hei   = 20    ' �֌W�Ȃ�
  'Fsize = 10
  'Fsize = 8

  xPos = C_XPOSL
  yPos = C_YPOSL
  wid  = C_WIDEL
  Fsize= C_FSIZE

  ' �����o���ꏊ, ��, �t�H���g�T�C�Y�̕ύX
  Call siz_MsgBox(sizFlg,xPos,wid,Fsize,xPos0,wid0,Fsize0)

  ' �X���C�h�֏����o��
  Call write_TOC(xPos,yPos,wid,hei,Fsize,nSids(),Conts(),Pagesc(),C_HIERARCHY,0, GFLAG)

End Sub

Public Sub removeListHierarchy()
  Dim SlideObj   As Slide
  Dim ShapeObj   As Shape
  Dim ShapeIndex As Integer

  For Each SlideObj In ActivePresentation.Slides          ' �e�X���C�h
    For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1   ' �e�V�F�C�v
      Set ShapeObj = SlideObj.Shapes(ShapeIndex)
      If ShapeObj.Type = msoTextBox Then           ' TextBox �ł���
        If ShapeObj.Name = C_HIERARCHY Then          ' ���O�� thierarchy �ł���.
          ShapeObj.Delete                          ' �I�u�W�F�N�g����������.
        End If
      End If
    Next ShapeIndex
  Next SlideObj
End Sub

' �����p�����̍쐬 sw =0 �ɂ����Ƃ�
'   �L�[���[�h���������ŕ\�����ꂽ�X���C�h
'   �w���p�ł͍폜���Ă���X���C�h�ɓ������������
' �w���p�����̍쐬: sw = 1 �ɂ����Ƃ�
'   �L�[���[�h���u�����ꂽ�X���C�h
'   �����ڂ��Ă���X���C�h���폜����
Public Sub flipTextForStudy(student As Integer, GFLAG As Integer)

  ' ���̃��[�`���ł͈ȉ��������ōs��:
  ' 1. �w���p��, �L�[���[�h���u��(_ ��)���ꂽ�X���C�h���쐬����.
  ' 2. ���K���̓������ڂ��Ă���X���C�h���폜����
  '
  ' �J���Ă��� powerpoint �t�@�C�� (hoge.pptm) �ɑ΂���,
  ' 1. �u�����, �폜�X���C�h�ԍ��������ꂽ�e�L�X�g�t�@�C��(default=hoge.txt) ��ǂݍ���.
  ' 2. ���� pptm(pptx) file(hoge.pptm)  -> hoge_for_student.pptm �ɕۑ����Ȃ���
  ' 3. ������u���ƃX���C�h�폜

  ' hoge.txt �̒��g:
  ' ::%%%::
  ' Entrez Gene    _____ ____     *
  ' SNP            ___            1,2,3-10,20
  ' db___          dbSNP          1,2,20
  ' ::@@@::
  ' 5,8,11
  ' hoge.txt �̒��g�I���
  '
  ' hoge.txt �� format:
  ' ::%%%::
  ' (�u�����镶����)\t\t\t\t(�u����̕�����)\t\t\t\t(�u������y�[�W)
  '    ��: �u������y�[�W�ɂ���:
  '        *: �S���̃y�[�W(���̏ꍇ�͏ȗ���)
  '        1,2,3: 1,2,3 �y�[�W
  '        3-10:  3 �y�[�W���� 10 �y�[�W�܂�
  ' ::@@@::
  ' (�y�[�W��),(�y�[�W��)-(�y�[�W��) ... <- �폜����y�[�W��
  ' �܂��͈ȉ��̂悤��, �ۑ�����y�[�W��������
  ' ::===::
  ' (�y�[�W��),(�y�[�W��)-(�y�[�W��) ... <- �c���y�[�W��

  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,GFLAG)

  ' text file �����݂��Ă��Ȃ���ΏI������
  If Len(Dir$(sNotesFilePath)) = 0 Then
    MsgBox (sNoteFilePath & " is missing")
    Exit Sub
  End If

  ' �ۑ�����w���p�����̃t�@�C����
  Dim sPptFileName As String
  If student = 1 Then
    spptFileName = newFileName("_for_student.pptm","do you want to save the file as ",GFLAG)
  Else
    spptFileName = newFileName("_for_teacher.pptm","do you want to save the file as ",GFLAG)
  End If

  ' �ʖ��ŏ㏑���ۑ�
  ' ��������͂��̃t�@�C���� Active �ƂȂ�̂Œ��ӂ���.
  Dim yvb As Integer
  If GFLAG = 1 Then
    yvb = MsgBox(ActivePresentation.Name & " will be saved as " & " " & sPptFileName, vbYesNo)
  Else
    yvb = vbYes
  End If
  'If MsgBox(ActivePresentation.Name & " will be saveed as " & " " & sPptFileName, vbYesNo) = vbNo Then
  If yvb = vbNo Then
    Exit Sub
  End If
  ActivePresentation.SaveAs (sCurrentFolder & sPptFileName)

  ' file reading
  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  Dim fString() As String ' �u����
  Dim tString() As String ' �u����
  Dim Pagesc()  As String ' �u��������y�[�W�� ��: Pagesc(1)=(1,2,3,5)
  Dim flag      As Integer
  flag = 0
  If Not IsNull(hBufs.Item(FLIP_MARK)) And Not hBufs.Item(FLIP_MARK) = "" Then
    Call bufRead_exList(hBufs.Item(FLIP_MARK),fString(),tString(),Pagesc(),C_TBB)
    flag = 1
  End If

  Dim nSids() As Integer
  Dim aSids() As Integer
  If Not IsNull(hBufs.Item(DELS_MARK)) And Not hBufs.Item(DELS_MARK) = "" Then
    Call bufRead_nList(hBufs.Item(DELS_MARK),nSids(),C_CMM)  ' vbTab -> ","
  Else
    Call bufRead_nList(hBufs.Item(ALIV_MARK),aSids(),C_CMM)  ' vbTab -> ","
  End If

  Dim npage As Integer
  ActivePresentation.Slides(1).Select
  npage = ActivePresentation.Slides.Count

  ' fString ����`����Ă���Βu������
  ' 2014/04/09
  If Not Sgn(fString) = 0 Then
    Dim fnum As Integer
    fnum = UBound(fString)

    ' �u������y�[�W
    ' doFlip(�u���y�AID(0-based), �X���C�hID(0-based))=1 -> �u������
    Dim vBuf()     As String
    Dim iBuf()     As String
    Dim doFlip()   As Integer

    If flag = 1 Then
      ' ���I�񎟌��z��̏�����
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

      ' ������u��
      Dim oSlide   As Slide
      Dim oShape   As Shape
      Dim Gry
      Gry = RGB(170,170,170)
      k = 1
      For Each oSlide In ActivePresentation.Slides   ' �e�X���C�h
        For Each oShape In oSlide.Shapes             ' �e�V�F�C�v
          ' MsgBox (oShape.Name & " " & oShape.Type & "+" & oShape.AutoShapeType)
          For j = 0 To fnum
            If Not IsNull(doFlip(j, k)) And doFlip(j, k) = 1 Then
              If student = 1 Then
                ' �u��������
                Call FindnRe(oShape, fString(j), tString(j))
              Else
                ' �u���������ɓ������ŏ���
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
  ' �����p slide(���͋󂢂Ă��邪�X���C�h�̍폜�͂���Ă��Ȃ�)���쐬.
  ' �X���C�h���폜����O�ɕʖ��ŕۑ����Ă���.
  'Dim sPptFileName_For_Teacher As String
  'sPptFileName_For_Teacher = Mid$(ActivePresentation.Name, 1, InStr(ActivePresentation.Name, ".") - 1) & "_for_teacher.pptm"
  'ActivePresentation.SaveAs (sCurrentFolder & sPptFileName_For_Teacher)

  Dim Gry2
  Gry2 = RGB(200,200,200)
  If student = 1 Then
    ' �w���p�ŕK�v�̂Ȃ��X���C�h�̍폜
    Call removeSlide(nSids(),aSids(),npage)
  Else
    Call watermarkSlide(nSids(),aSids(),npage, Gry2)
  End If

  ' 20131014
  ' �w���p�ɍ폜������� slide �� hoge_for_student.pptm �ɕۑ����Ă���.
  ' �������� �ۑ����܂��� �ƕ������̂��ʓ|�Ȃ̂�
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
    ' do nothing �m�F���Ȃ��Ői��
  Else
    ' �m�F����
    If MsgBox(message & "'" & sPptFileName & "'?", vbYesNo) = vbNo Then
      sPptFileName = InputBox("input file name", "�t�@�C���������", sPptFileName)
    End If
  End If
  newFileName = sPptFileName
End Function

Private Sub removeSlide(nSids() As Integer, aSids() As Integer, npage As Integer)
  ' 20130611
  ' i �ԖڃX���C�h(1-based) ��
  ' �폜����   dSld(i) = 1
  ' �폜���Ȃ� dSld(i) = 0
  Dim dSld() As Integer
  If Not Sgn(nSids) = 0 Then
    Call checkd_slde(nSids(), dSld(), npage)
  ElseIf Not Sgn(aSides) = 0 Then
    Call checkw_slde(aSids(), dSld(), npage)
  End If

  If Not Sgn(dSld) = 0 Then
    Dim i As Integer
    For i = npage To 1 Step -1 ' �X���C�h���폜����ƃX���C�h�ԍ����ς��̂�, ��납��v�Z����.
      If dSld(i) = 1 Then      ' �폜
        ActivePresentation.Slides(i).Delete
      End If
    Next i
  End If
End Sub

Private Sub watermarkSlide(nSids() As Integer, aSids() As Integer, npage As Integer, rgb As Variant)
  ' 20141114 from removeSlide
  ' i �ԖڃX���C�h(1-based) ��
  ' �������  dSld(i) = 1
  ' ���Ȃ�    dSld(i) = 0
  Dim dSld() As Integer
  If Not Sgn(nSids) = 0 Then
    Call checkd_slde(nSids(), dSld(), npage)
  ElseIf Not Sgn(aSides) = 0 Then
    Call checkw_slde(aSids(), dSld(), npage)
  End If

  If Not Sgn(dSld) = 0 Then
    Dim i As Integer
    For i = npage To 1 Step -1 ' �폜�͂��Ȃ��̂Ŕԍ��͕ς��Ȃ���, removeSlide �𓥏P����
      If dSld(i) = 1 Then      ' �������
        With ActivePresentation.Slides.Item(i).Shapes.AddLine(C_XPOSL,C_YPOSL,C_XPOSB,C_YPOSB)
          .Name               = "watermarks"
          .Line.ForeColor.RGB= rgb
          '.Line.DashStyle     =msoLineRoundDot
          .Line.Style = msoLineSingle
          .Line.Weight        =C_LWID
          .ZOrder(msoSendToBack) ' �ŉ��w�ɒu��
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
    If sForm <> "" Then  ' 2014/04/29 �y�[�W���������ꍇ�̒u��������͏����Ȃ�
      prnt = prnt & fString(k) & delm & tString(k) & delm & sForm & vbNewLine
    End If
  Next k

End Sub

Private Sub bufRead_exList(sBuf As Variant, fString() As String, tString() As String, Pagesc() As String,delm As String)
  Dim aBuf() As String

  aBuf = Split(sBuf,vbNewLine) ' ��s������

  Dim m As Integer
  m = UBound(aBuf) ' m = �s�� -1 (0-based)
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
        If k = 0 Then      ' from (�u����������)
          fString(i) = buf(j)
        ElseIf k = 1 Then ' to (�u���敶����)
          tString(i) = buf(j)
        ElseIf k = 2 Then ' page (�u������y�[�W�̃��X�g)
          sForm(i)  = buf(j) ' �Ƃ肠����������Ƃ��Ď��o��(��ŏ�������)
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
' page �̕\���̎d��
'
' sForm   = "1-3,4,6,10-12"    'text file
' iList() = (1,2,3,4,6,10,12)  ' �v�Z�@����
' sList   = "1,2,3,4,6,10,12"  ' �v�Z�@����(���ԑ�)
' (defList() = (1,2,3, .... npage))
' �̊Ԃ̑��ݕϊ� script
Private Sub sForm2iList(sForm As String, iList() As Integer, delm As String, defList() As Integer)
  Dim vBuf() As String
  vBuf = Split(sForm,delm)
  Dim j As Integer
  j = 0
  Dim p As Integer
  For p = 0 To UBound(vBuf)
    If vBuf(p) = "*" Then     '�S��
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
  ' �X���C�h�̖����𓾂�
  If pPt Is Nothing Then
    ActivePresentation.Slides(1).Select
    npage = ActivePresentation.Slides.Count
  Else
    npage = pPt.Slides.Count
  End If

  ReDim defList(npage-1)
  Dim i As Integer
  For i = 0 To npage-1
    defList(i)=i+1 ' �S���g���ꍇ�� Pages list (1,2,3,....npages)
  Next i
End Sub

'
' flip2
'
Private Sub FindnCol(oShape As Shape, fString As String, rgb As Variant)
  Dim i       As Integer
  ' shape �̃^�C�v�ɂ���ĕ���
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
      ElseIf oShape.HasSmartArt Then   ' placeholder ���� SmartArt ������ꍇ ��.
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
      Set oTxtRng = oShape.Table.Rows(iRows).Cells(iCols).Shape.TextFrame2.TextRange ' TextFrame2 ���g����.
      Set oTmpRng = oTxtRng.Find(fString)
      Do While ((Not oTmpRng = "") and (Not oTmpRng Is Nothing))
        oTmpRng.Font.Fill.ForeColor.RGB=rgb ' �F�̎w��̎d�����ς���Ă���
        oTmpRng.Font.Italic    = msoTrue
        oTmpRng.Font.UnderlineStyle=msoUnderlineWavyHeavyLine ' �g���������Ă���
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
  Set oTxtRng = oShape.TextFrame2.TextRange  ' TextFrame2 ���g����.
  Set oTmpRng = oTxtRng.Find(fString)
  'MsgBox oTmpRng
  ' TextFrame2 �ɂ���� oTmpRng �� "" �ɂȂ邱�Ƃ�������,
  ' Nothing �̏����������Ƃ��̂����ŃG���[�ƂȂ邱�Ƃ�����.
  ' �悭�킩��Ȃ��̂ŏ����𕡐��t���Ă���.
  ' Activpresentation.Slides(1).Shapes(1).TextFrame2.TextRange.Font.
  Do While ((Not oTmpRng = "") and (Not oTmpRng Is Nothing))
    oTmpRng.Font.Fill.ForeColor.RGB=rgb ' �F�̎w��̎d�����ς���Ă���. �ʓ|������.
    oTmpRng.Font.Italic    = msoTrue
    oTmpRng.Font.UnderlineStyle=msoUnderlineWavyHeavyLine ' �g���������Ă���
    Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
  Loop
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

Private Sub ColorText2(oShape As Shape, ByVal i As Integer, fString As String, rgb As Variant)
  Dim oTxtRng
  Dim oTmpRng
  Set oTxtRng = oShape.SmartArt.AllNodes(i).TextFrame2.TextRange ' ����͌��X TextFrame2 ������.
  Set oTmpRng = oTxtRng.Find(fString)
  Do While ((Not oTmpRng = "") and (Not oTmpRng Is Nothing)) ' Nothing �����ŃG���[�ɂȂ������Ɩ������ǔO�̂���
    oTmpRng.Font.Fill.ForeColor.RGB=rgb ' �F�̎w��̎d�����ς���Ă���.
    oTmpRng.Font.Italic    = msoTrue
    oTmpRng.Font.UnderlineStyle=msoUnderlineWavyHeavyLine ' �g���������Ă���
    Set oTmpRng = oTxtRng.Find(fString, After:=oTmpRng.Start + oTmpRng.Length)
  Loop
  Set oTxtRng = Nothing
  Set oTmpRng = Nothing
End Sub

'
' �����݂̂������ꍇ(TextFrame2 ���g��Ȃ��ėǂ�. �����S)
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
'         oTmpRng.Font.Underline = msoTrue  ' ����������
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
' flip �֘A
'
Private Sub FindnRe(oShape As Shape, fString As String, tString As String)
  Dim i       As Integer
  ' shape �̃^�C�v�ɂ���ĕ���
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
      ElseIf oShape.HasSmartArt Then   ' placeholder ���� SmartArt ������ꍇ ��.
        For i = 1 To oShape.SmartArt.AllNodes.Count
          Call ReplaceText2(oShape,i,fString, tString)
        Next
      End If
  End Select
  ' �������J��
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
  For Each oSlide In ActivePresentation.Slides   ' �e�X���C�h
    For Each oShape In oSlide.Shapes             ' �e�V�F�C�v
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
  ' ���������
  Set oSlide = Nothing
  Set oShape = Nothing
End Sub

Private Sub findSmartArt2(oDummy As Shape)
  Dim oSLide As Slide
  Dim oShape As Shape
  For Each oSlide In ActivePresentation.Slides   ' �e�X���C�h
    For Each oShape In oSlide.Shapes             ' �e�V�F�C�v
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
  ' ���������
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
'   ' shape �̃^�C�v�ɂ���ĕ���
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

  ' �ꏊ�w��
  Dim xPos, yPos, wid, hei
  xPos = 686
  yPos = 527

  wid = 40
  hei = 28.875

  ' font
  Dim Fname, Ename, Fsize
  Fname = "Arial"
  Ename = "�l�r �o�S�V�b�N"
  Fsize = 10

  ActivePresentation.Slides(1).Select     ' �ꖇ�ڂ̃X���C�h��I��
  Dim npage
  npage = ActivePresentation.Slides.Count ' �X���C�h�̑����𐔂���

  Dim i As Integer
  Dim Shp
  For i = 1 To npage
    ActivePresentation.Slides(i).Select
   ' i ���ڂ̃X���C�h
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
    Shp.Name = C_PAGENUM  ' textbox �ɖ��O�����Ă���
  Next
End Sub

'Public Sub pageCountsBox()
Private Sub pageCountsBox()
  ' �F�ݒ�
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

  ' �ꏊ�ݒ�
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
     ' �g��
      If (j Mod 10 = 0) Then
        oSp.Line.ForeColor.RGB=Red
      Else
        oSp.Line.ForeColor.RGB=Blk
      End If
     ' �h��Ԃ�
      If (j < i) Then
        oSp.Fill.ForeColor.RGB=Gry
      ElseIf (j = i) Then
        oSp.Fill.ForeColor.RGB=Col
      Else
        oSp.Fill.ForeColor.RGB=Wht
      End If
      ShapeList(j) = Shp.Name
    Next
    ' �O���[�v�����Ė��O������
    ActivePresentation.Slides(i).Shapes.Range(ShapeList()).Group.Name = C_COUNTBOX
  Next
  ActivePresentation.Slides(1).Select
  ' ���������
End Sub

Public Sub source_copy()
  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  '''Set hBufs = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)
  ' java file ���̎擾
  Dim jFiles() As String 
  jFiles= Split(hBufs.Item(SOUR_MARK),vbNewLine)

  Dim i As Integer
  For i= 0 To UBound(jFiles)
    Dim page As Integer
    page = i + 1
    Dim sContent As String
    Call fileContents(filePath(jFiles(i)),sContent)   ' �t�@�C���̒��g�𓾂�
    ActivePresentation.Slides.Add page, ppLayoutBlank ' �����X���C�h�̒ǉ�
    ActivePresentation.Slides(page).Select
    ' �ǉ������X���C�h�� textbox ���쐬
    Set Shp=ActivePresentation.Slides.Item(page).Shapes.AddTextbox(msoTextOrientationHorizontal,C_XPOSS,C_YPOSS,C_XWIDS,C_YWIDS)
    Shp.Select
    ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=page, Length:=0).Select
    With ActiveWindow.Selection.TextRange
      .Text = sContent ' �t�@�C���̒��g����������
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
    'Shp.Name = C_PAGENUM  ' textbox �ɖ��O�����Ă���
  Next i
End Sub

Public Sub write_txt_down_hierarchy()
  '
  ' ::|||:: �̏C��
  '
  ' �S�̂� hierarchy ���������
  '

  Dim npage As Integer
  npage = ActivePresentation.Slides.Count

  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  Dim aBuf() As String
  aBuf()=Split(hBufs.Item(GTOC_MARK),vbNewLine)

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
      hStr = aBuf(j) + vbNewLine + hStr ' �������Ȃ�
    ElseIf s = "*" Then
      hStr = "*" + aBuf(j) + vbNewLine + hStr ' * ������Ƃ��͈��������
      buf() = Split(aBuf(j),vbTab)
      page  = Val(buf(UBound(buf)))
    Else
      hStr = "*" + vbTab + aBuf(j) + vbTab + Trim(Str(page)) + vbNewLine + hStr' �����Ƃ��� *\t ������ 
    End If
  Next j

  ' GTOC_MARK key �̃f�[�^������
  hBufs.Remove(GTOC_MARK)
  ' GTOC_MARK key �̃f�[�^��V�����쐬
  hBufs.Add GTOC_MARK, hStr

  ' �㏑������_�ɂ��イ��
  If MsgBox("overwrite hierarchy to " &  sNotesFilePath & ": OK ?",vbYesNo) = vbYes Then
    Call hashPrint(hBufs,sNotesFilePath)
  Else
    MsgBox("do nothing")
  End If

End Sub

Public Sub write_txt_ListContents()

  '
  ' �K�w�ڎ� ::|||:: GTOC_MARK ����,
  ' ListContents �p�̖ڎ� ::###:: (�E��ɏ������e)
  ' �̃��X�g�������łƂ��Ă���
  ' ListContents �ŏ�������K�w���w�肷��
  '

  Dim npage As Integer
  npage = ActivePresentation.Slides.Count

  Dim maxH As Integer
  maxH = C_MAXH
  If MsgBox("the hire num of ListContents TOC: " & maxH & " ?",vbYesNo) = vbNo Then
    maxH = InputBox("the hierarchy number: ",maxH)
  End If

  ' text file(path) �̎擾
  Dim sNotesFilePath As String
  Dim sCurrentFolder As String
  Call notesFilePath(sNotesFilePath,sCurrentFolder,".txt",1,1)

  '''Set hBufs   = New Dictionary ' �ǂݍ��񂾕�����
  Dim hBufs As Object
  Set hBufs=CreateObject("Scripting.Dictionary")
  Call fileRead2(sNotesFilePath,hBufs)

  Dim aBuf() As String
  aBuf()=Split(hBufs.Item(GTOC_MARK),vbNewLine)

  Dim m As Integer
  m = UBound(aBuf)

  Dim j      As Integer
  Dim pnum() As Integer
  ReDim pnum(maxH)
  For j=0 To maxH
    pnum(j)=npage   ' �K�w 1 �̃y�[�W = pnum(1)
  Next j

  Dim hStr   As String
  Dim buf()  As String
  Dim star   As Integer
  For j=m To 0 Step -1  ' �Ōォ�猩�Ă���.
    If Left$(aBuf(j),1) = "*" Then
      buf()=Split(aBuf(j),C_TBB)
      star = Len(buf(0))
      If Not star > maxH Then
        If Not buf(UBound(buf)) = Trim(Str(pnum(star))) Then
          hStr = buf(1) + vbTab + buf(UBound(buf)) + "-" + Trim(Str(pnum(star))) + vbNewLine + hStr
        Else
          hStr = buf(1) + vbTab + buf(UBound(buf)) + vbNewLine + hStr ' 1 page �����̏ꍇ
        End If
        pnum(star)=Val(buf(UBound(buf)))-1
      End If
    End If
  Next j

  ' �㏑������_�ɂ��イ��
  If MsgBox("add ListContents to " &  sNotesFilePath & ": OK ?",vbYesNo) = vbYes Then
    ' TOCS_MARK key �̃f�[�^������
    If hBufs.Exists(TOCS_MARK) Then
      hBufs.Remove(TOCS_MARK)
    End If
    ' TOCS_MARK key �̃f�[�^��V�����쐬
    hBufs.Add TOCS_MARK, hStr
    Call hashPrint(hBufs,sNotesFilePath)
  Else
    MsgBox("do nothing")
  End If

End Sub

' ����������S�z���Ȃ��ėǂ��Ƃ�
'Private Sub fileContents(sNotesFilePath As String, sContent As String)
'  If Dir(sNotesFilePath) = "" Then  ' �t�@�C���̑��݂��m�F
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
'    sContent = sContent & sBuf & vbNewLine ' �����Ă��镶��������Ă���
'  Loop
'  Close iNotesFileNum
'End Sub

' UTF-8 �e�L�X�g�t�@�C����ǂݍ��ނƂ�
Private Sub fileContents(sNotesFilePath As String, sContent As String)
  If Dir(sNotesFilePath) = "" Then  ' �t�@�C���̑��݂��m�F
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
  Fname = Replace(Fname,"/","\")
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
    ' ����Fname ���t�@�C���������̏ꍇ��, �J�����g�f�B���N�g���ɂ���Ƃ݂Ȃ���
    ' ��΃p�X�ɂ���.
    '���邢�� "." ����n�܂�ꍇ�ɂ͑��΃p�X�ŏ����Ă���ƍl����
    ' ������f�B���N�g������ǋL���Đ�΃p�X�ɂ���
    ' ����ȊO�̏ꍇ��, ��΃p�X�ŏ����Ă���Ƃ��񂪂���.
    If InStr(Fname, "\") = 0 Then
      Fname = ActivePresentation.Path & "\" & Fname
      'ElseIf Left$(Fname,2) = ".\" Then   ' 2014/04/11 �폜
      '  Fname = ActivePresentation.Path & "\" & Mid(Fname,2)
    ElseIf Left(Fname,1) = "." Then ' 2014/04/11 �悭�킩��Ȃ�.
      Fname = ActivePresentation.Path & "\" & Fname
    End If
  End If
  filePath=Fname
End Function