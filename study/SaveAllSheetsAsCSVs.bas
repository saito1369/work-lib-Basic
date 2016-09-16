Attribute VB_Name = "SaveAllSheetsAsCSVs"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub saveAsCSV()
  Dim delm As String
  Dim os   As Integer
  os = 0
  delm = "\"
  If Application.OperatingSystem Like "Macintosh*" Then
    delm = ":"
    os   = 1   ' flag for mac
  End If

  
  Dim mysheet As Worksheet
  Dim fname As String
  Dim path As String

  Application.DisplayAlerts = False

  For Each mysheet In ActiveWorkbook.Worksheets
    fname = mysheet.Name & ".csv"
    path = ActiveWorkbook.path & delm & fname
    mysheet.Activate
    mysheet.Copy
    'MsgBox fname
    ActiveWorkbook.SaveAs Filename:=fname, _
                  FileFormat:=xlCSV, CreateBackup:=False
    Workbooks(fname).Close
  Next
  
  Application.DisplayAlerts = True
  
End Sub
