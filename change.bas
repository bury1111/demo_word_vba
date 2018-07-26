Attribute VB_Name = "NewMacros"
Sub main()
    Call change
    Call delete_lab
End Sub

Sub change()
     Dim i As Paragraph, mt, oRang As Range, n%, m%
    With CreateObject("vbscript.regexp")
        .Pattern = "(\[add\](.*?)\[/add\])"
        .Global = True: .IgnoreCase = False: .MultiLine = True
        For Each i In ActiveDocument.Paragraphs
            For Each mt In .Execute(i.Range.Text)
                m = mt.FirstIndex: n = mt.Length
                Set oRang = ActiveDocument.Range(i.Range.Start + m, i.Range.Start + m + n)
                'MsgBox (oRang)
                oRang.Font.Color = wdColorBlue
                oRang.Font.Underline = wdUnderlineSingle
            Next
        Next
    End With
    
    With CreateObject("vbscript.regexp")
        .Pattern = "(\[del\](.*?)\[/del\])"
        .Global = True: .IgnoreCase = False: .MultiLine = True
        For Each i In ActiveDocument.Paragraphs
            For Each mt In .Execute(i.Range.Text)
                m = mt.FirstIndex: n = mt.Length
                Set oRang = ActiveDocument.Range(i.Range.Start + m, i.Range.Start + m + n)
                'MsgBox (oRang)
                oRang.Font.Color = wdColorRed
                oRang.Font.StrikeThrough = True
            Next
        Next
    End With
End Sub
'É¾³ý±êÇ©
Sub delete_lab()

Selection.WholeStory

    With Selection.Find
    
        .ClearFormatting
        
        .Replacement.ClearFormatting
        
        .Text = "\[*\]"
        
        .Replacement.Text = ""
        
        .Forward = True
        
        .Execute Replace:=wdReplaceAll
        
    End With
    
End Sub

