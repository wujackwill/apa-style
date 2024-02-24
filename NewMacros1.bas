Attribute VB_Name = "NewMacros1"

Sub AdjustStyles()
    Dim doc As Document
    Set doc = ActiveDocument ' 使用当前活动文档

    Dim para As Paragraph
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    For Each para In doc.Paragraphs
        ' 使用正则表达式检查是否包含中文句子格式 "外语教学与研究, 52(1)"
        regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+\(\d+\).*"
        If regex.Test(para.Range.text) Then
            ' 中文段落设置字体为宋体
            
            ' 清除中文段落的斜体
            para.Range.Font.Italic = False
        End If
    Next para
    
    ' 遍历文档中的每一个段落
    For Each para In doc.Paragraphs
        ' 使用正则表达式检查是否包含中文句子格式 ", 外语教学与研究, 52(1),"
        regex.Pattern = "([.,]\s*)([一-龥\s]+)([,])"

        ' 查找匹配项
        Dim matches As Object
        Set matches = regex.Execute(para.Range.text)
        
        ' 处理匹配项
        For Each match In matches
            ' 获取匹配项的范围
            Dim matchRange As Range
            Set matchRange = para.Range.Duplicate
            matchRange.SetRange para.Range.Start + match.FirstIndex, para.Range.Start + match.FirstIndex + match.Length - 1

            ' 设置中文段落的斜体
            matchRange.Font.Italic = True
        Next match
    Next para
    

End Sub


