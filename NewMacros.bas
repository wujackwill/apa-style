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
            
            ' Çå³ıÖĞÎÄ¶ÎÂäµÄĞ±Ìå
            para.Range.Font.Italic = False
        End If
    Next para
End Sub

