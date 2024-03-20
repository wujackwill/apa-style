Attribute VB_Name = "NewMacros1"
Sub Adjuststyles()

    Call italicstyles
    
    Call RemoveParagraphIndentationAndDoubleSpacing
End Sub


Public Sub italicstyles()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim para As Paragraph
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+.*"

    For Each para In doc.Paragraphs
        If regex.Test(para.Range.text) Then
            para.Range.Font.Italic = False
        End If
    Next para

    regex.Pattern = "([\u4e00-\u9fa5]+), \d+.*"

    For Each para In doc.Paragraphs
        Dim matches As Object
        Set matches = regex.Execute(para.Range.text)
        
        For Each match In matches
            Dim matchRange As Range
            Set matchRange = para.Range.Duplicate
            matchRange.Find.text = match.SubMatches(0)
            matchRange.Find.Execute
            matchRange.Font.Italic = True
        Next match
    Next para
End Sub


Public Sub RemoveParagraphIndentationAndDoubleSpacing()
    Dim para As Paragraph
    Dim text As String
    Dim i As Long
    
    For Each para In ActiveDocument.Paragraphs
        text = para.Range.text
        If Left(text, 1) = "[" Then
            para.LeftIndent = 0
            para.FirstLineIndent = 0
            para.LineSpacingRule = wdLineSpaceSingle
        End If
    Next para
End Sub


