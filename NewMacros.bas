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
            para.Range.Font.Name = "����"
            
            ' ������Ķ����б��
            para.Range.Font.Italic = False
        Else
            ' ���������е�ÿһ�� Run
            For Each run In para.Range.Words
                ' ��鵱ǰ Run ���ı��Ƿ�Ϊ����������
                If IsArabicNumber(run.text) Then
                    ' ���ð��������� Run ������Ϊ Times New Roman
                    run.Font.Name = "Times New Roman"
                End If
            Next run
        End If
    Next para
End Sub

Function IsArabicNumber(text As String) As Boolean
    ' ʹ��������ʽ����ı��Ƿ�Ϊ����������
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' ƥ�䰢�������ֵ�������ʽ
    regex.Pattern = "^\d+(\(\d+\))?(\.\d+)?$"
    
    ' ����ı��Ƿ����������ʽ
    IsArabicNumber = regex.Test(text)
End Function



