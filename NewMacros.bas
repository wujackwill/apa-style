Attribute VB_Name = "NewMacros"
Sub ZoteroAddEditBibliography()
'
' ZoteroAddEditBibliography å®
'
'

Sub AdjustStyles()
    Dim doc As Document
    Set doc = ActiveDocument ' Ê¹ÓÃµ±Ç°»î¶¯ÎÄµµ

    Dim para As Paragraph
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    For Each para In doc.Paragraphs
        ' Ê¹ÓÃÕıÔò±í´ïÊ½¼ì²éÊÇ·ñ°üº¬ÖĞÎÄ¾ä×Ó¸ñÊ½ "ÍâÓï½ÌÑ§ÓëÑĞ¾¿, 52(1)"
        regex.Pattern = ".*[\u4e00-\u9fa5]+, \d+\(\d+\).*"
        If regex.Test(para.Range.text) Then
            ' ÖĞÎÄ¶ÎÂäÉèÖÃ×ÖÌåÎªËÎÌå
            
            ' Çå³ıÖĞÎÄ¶ÎÂäµÄĞ±Ìå
            para.Range.Font.Italic = False
        End If
    Next para
    
    ' ±éÀúÎÄµµÖĞµÄÃ¿Ò»¸ö¶ÎÂä
    For Each para In doc.Paragraphs
        ' Ê¹ÓÃÕıÔò±í´ïÊ½¼ì²éÊÇ·ñ°üº¬ÖĞÎÄ¾ä×Ó¸ñÊ½ ", ÍâÓï½ÌÑ§ÓëÑĞ¾¿, 52(1),"
        regex.Pattern = "([.,]\s*)([Ò»-ı›\s]+)([,])"

        ' ²éÕÒÆ¥ÅäÏî
        Dim matches As Object
        Set matches = regex.Execute(para.Range.text)
        
        ' ´¦ÀíÆ¥ÅäÏî
        For Each match In matches
            ' »ñÈ¡Æ¥ÅäÏîµÄ·¶Î§
            Dim matchRange As Range
            Set matchRange = para.Range.Duplicate
            matchRange.SetRange para.Range.Start + match.FirstIndex, para.Range.Start + match.FirstIndex + match.Length - 1

            ' ÉèÖÃÖĞÎÄ¶ÎÂäµÄĞ±Ìå
            matchRange.Font.Italic = True
        Next match
    Next para
    

End Sub


