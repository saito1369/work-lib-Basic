Attribute VB_Name = "TiledFig"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub addPictures()

  ' ������ powerpoint �����ɐ}�𖄂ߍ���
  ' �ݒ�t�@�C����ǂݍ���
  ' powerpoint �̃y�[�W��\t�}�̃t�@�C����\ttop�ʒu(cm)\tleft �ʒu(cm)\t�{��\n
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
  Dim sCurrentFolder As String   ' ppt ������ folder
  Dim sNotesFileName As String   ' text file ��
  sCurrentFolder = ActivePresentation.Path & delm
  sNotesFileName = Mid$(ActivePresentation.Name, 1, InStr(ActivePresentation.Name, ".") - 1)
  sNotesFileName = sNotesFileName & ".txt"
  If MsgBox("do you want to make the table of contents using the file '" & sNotesFileName & "'?", vbYesNo) = vbNo Then
    sNotesFileName = InputBox("input file name", "�}�}���ݒ�t�@�C�����̓���", sNotesFileName)
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

  Dim Pages()  As Integer ' powerpoint �̃y�[�W��(1-based)
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
    If left(sBuf,1) = "#" Then    '# �Ŏn�܂��Ă���s�̓R�����g
      ' do nothing
    ElseIf Len(sBuf) > 0 Then    ' ��s�łȂ���Γǂݍ���
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
        Fname = Replace(Fname,"\",":")
        ' �t�@�C�������������Ă���ꍇ��, �J�����g�f�B���N�g�����Ɖ��肷��.
        If InStr(Fname, ":") = 0 Then Fname = sCurrentFolder & Fname
      Else 'win
        ' ��{�I�ɂ�, �t�H���_��؂�� windows �`���ŏ������Ƃɂ���.
        'Fname = Replace(Fname, ":","\")
        If InStr(Fname, "\") = 0 Then Fname = sCurrentFolder & Fname
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
    
    ' slide ��������Βǉ����� 2015/04/16
    Dim SlideObj As Slide
    If Pages(k) <= Pptt.Slides.Count Then  ' Pptt.Slides.Count = �X���C�h�̖���
      Set SlideObj = Pptt.Slides(Pages(k)) ' Slide �����ɂ���ꍇ
    Else
      Set SlideObj = Pptt.Slides.Add(Pages(k),ppLayoutBlank) ' Slide ��V���ɒǉ�����ꍇ
    End If
    Set ShapeObj = SlideObj.Shapes.addPicture(Fnames(k),msoFalse,msoTrue,pleft,ptop,-1,-1)
    With ShapeObj
      .ScaleHeight ratios(k), msoTrue
      .ScaleWidth  ratios(k), msoTrue
      .ZOrder      msoSendToBack        ' �Ŕw�ʂɈړ�
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
      'If ShapeObj.Type = msoGroup Then            ' �O���[�v
        If ShapeObj.Name = "addedpicture" Then      ' ���O�� "addedpicture"
          ShapeObj.Delete
        End If
      'End If
    Next ShapeIndex
  Next SlideObj
End Sub
