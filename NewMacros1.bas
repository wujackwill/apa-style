Attribute VB_Name = "NewMacros1"

Sub AdjustStyles()
    Dim doc As Document
    Set doc = ActiveDocument ' ʹ�õ�ǰ��ĵ�

    Dim para As Paragraph
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    For Each para In doc.Paragraphs
        ' ʹ��������ʽ����Ƿ�������ľ��Ӹ�ʽ "�����ѧ���о�, 52(1)"
        regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+\(\d+\).*"
        If regex.Test(para.Range.text) Then
            ' ���Ķ�����������Ϊ����
            
            ' ������Ķ����б��
            para.Range.Font.Italic = False
        End If
    Next para
    
    ' �����ĵ��е�ÿһ������
    For Each para In doc.Paragraphs
        ' ʹ��������ʽ����Ƿ�������ľ��Ӹ�ʽ ", �����ѧ���о�, 52(1),"
        regex.Pattern = "([.,]\s*)([һ-��\s]+)([,])"

        ' ����ƥ����
        Dim matches As Object
        Set matches = regex.Execute(para.Range.text)
        
        ' ����ƥ����
        For Each match In matches
            ' ��ȡƥ����ķ�Χ
            Dim matchRange As Range
            Set matchRange = para.Range.Duplicate
            matchRange.SetRange para.Range.Start + match.FirstIndex, para.Range.Start + match.FirstIndex + match.Length - 1

            ' �������Ķ����б��
            matchRange.Font.Italic = True
        Next match
    Next para
    

End Sub


