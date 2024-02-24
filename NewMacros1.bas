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

    ' 使用正则表达式检查是否包含中文句子格式 "外语教学与研究, 52(1)"
    regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+\(\d+\).*"

    ' 处理匹配项
    For Each para In doc.Paragraphs
        If regex.Test(para.Range.text) Then
            para.Range.Font.Italic = False
        End If
    Next para

    ' 再把中文斜体
    regex.Pattern = "([.,]\s*)([一-龥\s]+)([,])"

    ' 处理匹配项
    For Each para In doc.Paragraphs
        Dim matches As Object
        Set matches = regex.Execute(para.Range.text)
        
        For Each match In matches
            Dim matchRange As Range
            Set matchRange = para.Range.Duplicate
            matchRange.SetRange para.Range.Start + match.FirstIndex, para.Range.Start + match.FirstIndex + match.Length - 1
            matchRange.Font.Italic = True
        Next match
    Next para
End Sub

