Attribute VB_Name = "Module1"
Sub a()
  
  If IsNumeric(Range("A1")) = True Then
    If Sgn(Range("A1")) = -1 Then
      
      MsgBox "���̐�����[��"
      
    ElseIf IsNumeric(Range("A1")) = True And Sgn(Range("A1")) = 1 Then
        
      MsgBox "������[��"
        
    End If
    
  ElseIf IsNumeric(Range("A1")) = False Then
      MsgBox "����"
      
  End If


End Sub
