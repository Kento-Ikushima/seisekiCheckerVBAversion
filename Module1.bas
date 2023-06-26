Attribute VB_Name = "Module1"
Sub a()
  
  If IsNumeric(Range("A1")) = True Then
    If Sgn(Range("A1")) = -1 Then
      
      MsgBox "•‰‚Ì”‚¾‚æ[‚ñ"
      
    ElseIf IsNumeric(Range("A1")) = True And Sgn(Range("A1")) = 1 Then
        
      MsgBox "³‚¾‚æ[‚ñ"
        
    End If
    
  ElseIf IsNumeric(Range("A1")) = False Then
      MsgBox "‚ ‚Ù"
      
  End If


End Sub
