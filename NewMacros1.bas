Attribute VB_Name = "NewMacros1"
Sub AdjustStyles()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim para As Paragraph
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 使用正则表达式检查是否包含中文句子格式 "外语教学与研究, 52"
    regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+.*"

    ' 处理匹配项
    For Each para In doc.Paragraphs
        If regex.Test(para.Range.text) Then
            para.Range.Font.Italic = False
        End If
    Next para

    ' 使用正则表达式匹配中文文本，不论后面的数字格式如何，只要逗号之间夹着数字就匹配成功
    regex.Pattern = "([\u4e00-\u9fa5]+), \d+.*"

    ' 处理匹配项
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

