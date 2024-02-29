Attribute VB_Name = "NewMacros1"
Sub AdjustStyles()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim para As Paragraph
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' ʹ��������ʽ����Ƿ�������ľ��Ӹ�ʽ "�����ѧ���о�, 52"
    regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+.*"

    ' ����ƥ����
    For Each para In doc.Paragraphs
        If regex.Test(para.Range.text) Then
            para.Range.Font.Italic = False
        End If
    Next para

    ' ʹ��������ʽƥ�������ı������ۺ�������ָ�ʽ��Σ�ֻҪ����֮��������־�ƥ��ɹ�
    regex.Pattern = "([\u4e00-\u9fa5]+), \d+.*"

    ' ����ƥ����
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

