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

'add��ǩ�� ���ݸı�
Sub add_change()
    
    With Selection.Find
        
        .MatchWildcards = True
        
        .Text = "(add)(*)(/add)"
        
        .Wrap = wdFindContinue
        
        Do While .Execute
    
     If .Found = True Then
        
         ActiveDocument.Range(Selection.Start + 4, Selection.End - 5).Select
        
         'MsgBox Selection.Text   '����ѡ������
        
          Selection.Font.Color = wdColorBlue
         
          Selection.Font.Underline = wdUnderlineSingle
        
        End If
    
    Loop

    End With
        
End Sub


'del��ǩ�� ���ݸı�
    Sub del_change()
    
    With Selection.Find
        
        .MatchWildcards = True
        
        .Text = "(del)(*)(/del)"
        
        .Wrap = wdFindContinue
        
       Do While .Execute
    
    If .Found = True Then
        
        ActiveDocument.Range(Selection.Start + 4, Selection.End - 5).Select
        
       'MsgBox Selection.Text   '����ѡ������
        
         Selection.Font.Color = wdColorRed
        
        Selection.Font.StrikeThrough = True
        
        End If

    Loop
    
    End With

    End Sub

'ɾ����ǩ
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
