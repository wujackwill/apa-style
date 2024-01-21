Attribute VB_Name = "NewMacros"
Sub ZoteroAddEditBibliography()
'
' ZoteroAddEditBibliography å®
'
'
Sub AdjustStles()
    Dim doc As Document
    Set doc = ActiveDocument ' Ê¹ÓÃµ±Ç°»î¶¯ÎÄµµ

    Dim para As Paragraph
    Dim run As Range
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' ±éÀúÎÄµµÖĞµÄÃ¿Ò»¸ö¶ÎÂä
    For Each para In doc.Paragraphs
        ' Ê¹ÓÃÕıÔò±í´ïÊ½¼ì²éÊÇ·ñ°üº¬ÖĞÎÄ¾ä×Ó¸ñÊ½ "ÍâÓï½ÌÑ§ÓëÑĞ¾¿, 52(1)"
        regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+\(\d+\).*"
        If regex.Test(para.Range.text) Then
            ' ÖĞÎÄ¶ÎÂäÉèÖÃ×ÖÌåÎªËÎÌå
            para.Range.Font.Name = "ËÎÌå"
            
            ' Çå³ıÖĞÎÄ¶ÎÂäµÄĞ±Ìå
            para.Range.Font.Italic = False
        Else
            ' ±éÀú¶ÎÂäÖĞµÄÃ¿Ò»¸ö Run
            For Each run In para.Range.Words
                ' ¼ì²éµ±Ç° Run µÄÎÄ±¾ÊÇ·ñÎª°¢À­²®Êı×Ö
                If IsArabicNumber(run.text) Then
                    ' ÉèÖÃ°¢À­²®Êı×Ö Run µÄ×ÖÌåÎª Times New Roman
                    run.Font.Name = "Times New Roman"
                End If
            Next run
        End If
    Next para
End Sub

Function IsArabicNumber(text As String) As Boolean
    ' Ê¹ÓÃÕıÔò±í´ïÊ½¼ì²éÎÄ±¾ÊÇ·ñÎª°¢À­²®Êı×Ö
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Æ¥Åä°¢À­²®Êı×ÖµÄÕıÔò±í´ïÊ½
    regex.Pattern = "^\d+(\(\d+\))?(\.\d+)?$"
    
    ' ¼ì²éÎÄ±¾ÊÇ·ñ·ûºÏÕıÔò±í´ïÊ½
    IsArabicNumber = regex.Test(text)
End Function



