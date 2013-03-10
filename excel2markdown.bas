' Excel2Markdown
' Author: Tatsuya Hoshino
' Update: 2013/03/10

Attribute VB_Name = "Excel2Markdown"
Option Explicit

Const editMenuItemIndex As Integer = 2               ' Edit menu item index
Const newMenuTitle As String = "Copy As Markdown"    ' New menu title
Const newMenuTag   As String = "Excel2MarkdownAddin" ' New menu tag

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

  If rng.Font.Underline <> xlUnderlineStyleNone Then
    cellStr = "<ins>" & cellStr & "</ins>"
  End If
  If rng.Font.Strikethrough Then
    cellStr = "<del>" & cellStr & "</del>"
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

Sub Auto_Open()
  Dim editMenuItem As CommandBarPopup, newMenu As CommandBarControl
  Set editMenuItem = CommandBars.ActiveMenuBar.Controls(editMenuItemIndex)
  Set newMenu = editMenuItem.Controls.Add(Type:=msoControlButton, Temporary:=True)

  With newMenu
    .Caption  = newMenuTitle
    .onAction = "Excel2Markdown.Excel2Markdown"
    .Tag      = newMenuTag
  End With
End Sub

Sub Auto_Close()
  Dim editMenuItem As CommandBarPopup
  Set editMenuItem = CommandBars.ActiveMenuBar.Controls(editMenuItemIndex)

  editMenuItem.Controls(newMenuTitle).Delete
End Sub
