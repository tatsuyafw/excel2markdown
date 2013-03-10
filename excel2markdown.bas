' excel2markdown
' Author: Tatsuya Hoshino
' Update: 2013/03/10

Attribute VB_Name = "excel2markdown"

Function CopyCellAsMarkdown(row As Integer, col As Integer) As String
  Dim cellStr As String
  Dim rng As Range
  Set rng = Selection.Cells(row, col)
  cellStr = Selection.Cells(row, col)

  If rng.Font.Bold Then
    cellStr = "**" & cellStr & "**"
  ElseIf rng.Font.Italic Then
    cellStr = "*"  & cellStr & "*"
  End If

  CopyCellAsMarkdown = cellStr
End Function

Function ReadLine(row As Integer) As String
  Dim strLine As String
  Dim col As Integer
  strLine = "|"

  For col = 1 To Selection.Columns.Count
    strLine = strLine & CopyCellAsMarkdown(row, col) & "|"
  Next

  ReadLine = strLine
End Function

Function HeaderLine(col As Integer) As String
  Dim strLine As String
  Dim x As Integer
  strLine = "|"

  For x = 1 To col
    strLine = strLine & "---" & "|"
  Next

  HeaderLine = strLine
End Function

' main
Sub Excel2Markdown()
  Dim table As String
  Dim row As Integer, col As Integer
  Dim CB As New DataObject

  ' read table header
  table = ReadLine(1) & vbNewLine

  ' add header line
  table = table & HeaderLine(Selection.Columns.Count) & vbNewLine

  ' read table body
  For row = 2 To Selection.Rows.Count
    table = table & ReadLine(row) & vbNewLine
  Next

  ' copy table to clip board
  With CB
    .SetText table
    .PutInClipboard
  End With
End Sub
