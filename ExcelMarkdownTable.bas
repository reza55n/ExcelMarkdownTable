Sub excelMarkdownTable()
    Dim R As Range, i As Long, j  As Long, Buffer As String
    
    Set R = Selection
    Buffer = ""

    If R.Cells.Count > 500 Then
        If MsgBox("You are going to convert " & R.Cells.Count & " cells that may take a longer time." & _
                vbNewLine & "Do you want to continue?", vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If


    For j = 1 To R.Columns.Count
        Buffer = Buffer & "| " & Replace(R.Rows(1).Cells(j), "|", "\|") & " "
    Next
    Buffer = Buffer & "|" & vbNewLine


    For j = 1 To R.Columns.Count
        Select Case R.Rows(1).Cells(j).HorizontalAlignment
            Case xlLeft
                Buffer = Buffer & "| :--- "
            Case xlRight
                Buffer = Buffer & "| ---: "
            Case xlCenter
                Buffer = Buffer & "| :---: "
        End Select
    Next
    Buffer = Buffer & "|" & vbNewLine
    
    
    For i = 2 To R.Rows.Count
        For j = 1 To R.Columns.Count
            Buffer = Buffer & "| " & Replace(R.Rows(i).Cells(j), "|", "\|") & " "
        Next
        Buffer = Buffer & "|" & vbNewLine
    Next
    
    Clipboard Buffer
End Sub


Function Clipboard(Optional StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim x As Variant

'Store as variant for 64-bit VBA support
x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With

End Function
