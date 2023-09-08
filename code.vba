Sub DeleteDuplicates()
    Dim ws As Worksheet
    Dim rng As Range
    
    ' Set the worksheet and range where you want to remove duplicates
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set rng = ws.Range("A1:A10") ' Update the range as per your requirements
    
    ' Remove duplicates
    rng.RemoveDuplicates Columns:=Array(1), Header:=xlNo
    
    ' Optionally, you can also specify the Header argument as xlYes if your range has a header row
    
    MsgBox "Duplicates removed successfully."
End Sub
