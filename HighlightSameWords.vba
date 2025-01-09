Sub HighlightSameWords()
    Dim selectedText As String
    Dim currentRange As Range
    Dim isHighlighted As Boolean

    ' Get the selected text or word at cursor
    If Selection.Type = wdSelectionNormal Then
        If Selection.Text = "" Then
            ' No selection - use word at cursor
            selectedText = Selection.Words(1).Text
        Else
            ' Use selected text
            selectedText = Selection.Text
        End If

        selectedText = Trim(selectedText)

        ' Remove any trailing punctuation
        If Right(selectedText, 1) Like "[!a-zA-Z0-9]" Then
            selectedText = Left(selectedText, Len(selectedText) - 1)
        End If

        ' Only proceed if there's actual text
        If Len(selectedText) > 0 Then
            ' Check if the current word is already highlighted
            isHighlighted = Selection.Range.HighlightColorIndex = wdYellow

            ' Clear existing highlighting
            ActiveDocument.Range.HighlightColorIndex = wdNoHighlight

            ' If it wasn't highlighted before, highlight all instances
            If Not isHighlighted Then
                ' Create a range for the whole document
                Set currentRange = ActiveDocument.Range

                ' Find and highlight all instances
                With currentRange.Find
                    .Text = selectedText
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = True
                    .MatchWildcards = False
                    .Forward = True

                    ' Find first match
                    Do While .Execute
                        ' Highlight the found text
                        currentRange.HighlightColorIndex = wdYellow

                        ' Move the range start to after the current match
                        currentRange.Start = currentRange.End

                        ' If we've reached the end of the document, stop
                        If currentRange.Start >= ActiveDocument.Range.End Then
                            Exit Do
                        End If
                    Loop
                End With
            End If
        End If
    End If
End Sub
