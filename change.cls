VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub change()
     Dim i As Paragraph, mt, oRang As Range, n%, m%
Set regx = CreateObject("vbscript.regexp")
        regx.Pattern = "(\[add\](.*?)\[/add\])"
        regx.Global = True
        regx.IgnoreCase = False
        regx.MultiLine = True
  
        For Each i In ActiveDocument.Paragraphs
            For Each mt In .Execute(i.Range.Text)
                m = mt.FirstIndex: n = mt.Length
                Set oRang = ActiveDocument.Range(i.Range.Start + m, i.Range.Start + m + n)
                'MsgBox (oRang)
                oRang.Font.Color = wdColorBlue
                oRang.Font.Underline = wdUnderlineSingle
            Next
        Next

         regx.Pattern = "(\[del\](.*?)\[/del\])"
        For Each i In ActiveDocument.Paragraphs
            For Each mt In .Execute(i.Range.Text)
                m = mt.FirstIndex: n = mt.Length
                Set oRang = ActiveDocument.Range(i.Range.Start + m, i.Range.Start + m + n)
                'MsgBox (oRang)
                oRang.Font.Color = wdColorRed
                oRang.Font.StrikeThrough = True
            Next
        Next
   
End Sub

