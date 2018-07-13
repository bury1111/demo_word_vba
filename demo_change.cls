VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub main()
    Call add_change
    Call del_change
    Call delete_lab
End Sub

'add标签内 内容改变
Sub add_change()
    
    With Selection.Find
        
        .MatchWildcards = True
        
        .Text = "(add)(*)(/add)"
        
        .Wrap = wdFindContinue
        
        Do While .Execute
    
     If .Found = True Then
        
         ActiveDocument.Range(Selection.Start + 4, Selection.End - 5).Select
        
         'MsgBox Selection.Text   '测试选中内容
        
          Selection.Font.Color = wdColorBlue
         
          Selection.Font.Underline = wdUnderlineSingle
        
        End If
    
    Loop

    End With
        
End Sub


'del标签内 内容改变
    Sub del_change()
    
    With Selection.Find
        
        .MatchWildcards = True
        
        .Text = "(del)(*)(/del)"
        
        .Wrap = wdFindContinue
        
       Do While .Execute
    
    If .Found = True Then
        
        ActiveDocument.Range(Selection.Start + 4, Selection.End - 5).Select
        
       'MsgBox Selection.Text   '测试选中内容
        
         Selection.Font.Color = wdColorRed
        
        Selection.Font.StrikeThrough = True
        
        End If

    Loop
    
    End With

    End Sub

'删除标签
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
