Public Sub MyLoops()
  Dim i as Integer
  i=1
  Do While i<=0 
    If ActiveCell.Value>10 Then
      ActiveCell.Interior.color=RGT(255,0,0)
    End If
    ActiveCell.Offset(1,0).Select
    i=i+1
  Loop
End Sub
