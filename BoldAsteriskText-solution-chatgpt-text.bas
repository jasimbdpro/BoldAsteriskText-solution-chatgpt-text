Attribute VB_Name = "BoldAsteriskText"
Sub BoldAsteriskText()
    Dim doc As Document
    Dim rng As Range
    Dim findRng As Range
    Dim startPos As Long
    Dim endPos As Long
    
    Set doc = ActiveDocument
    Set rng = doc.Content
    
    With rng.Find
        .Text = "\*\*[!\*]@\*\*"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        
        Do While .Execute
            ' Set the range to the found text, excluding the asterisks
            startPos = rng.Start + 2
            endPos = rng.End - 2
            
            ' Check if the range is valid
            If startPos < endPos Then
                ' Make the text between the asterisks bold
                Set findRng = doc.Range(Start:=startPos, End:=endPos)
                findRng.Font.Bold = True
                
                ' Remove the asterisks
                doc.Range(rng.Start, rng.Start + 2).Text = ""
                doc.Range(rng.End - 2, rng.End).Text = ""
            End If
            
            ' Collapse the range to move to the next instance
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

