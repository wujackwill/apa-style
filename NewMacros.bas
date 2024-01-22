Attribute VB_Name = "NewMacros"
Sub ZoteroAddEditBibliography()
'
' ZoteroAddEditBibliography 宏
'
'
Sub AdjustStles()
    Dim doc As Document
    Set doc = ActiveDocument ' ʹ�õ�ǰ��ĵ�

    Dim para As Paragraph
    Dim run As Range
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' �����ĵ��е�ÿһ������
    For Each para In doc.Paragraphs
        ' ʹ��������ʽ����Ƿ�������ľ��Ӹ�ʽ "�����ѧ���о�, 52(1)"
        regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+\(\d+\).*"
        If regex.Test(para.Range.text) Then
            ' ���Ķ�����������Ϊ����
            
            ' ������Ķ����б��
            para.Range.Font.Italic = False
        End If
    Next para
End Sub

