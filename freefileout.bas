Sub freefileout()
'This procedure uses CATIA as the host application.
'Loop through the open documents and collect their name and index. Output this data to a CSV.
'Next steps: save the file to C:\temp with a unique filename to easily collect different feature trees from multiple documents.
'            categorize the parts based on lamelle, motif, creation lam, etc.
 
'Open a new freefile in C:\Temp to capture the data.
OutputFile = "C:\Temp\structure.txt"
fnum1 = FreeFile()

'Instantiate objects
Dim objProd As Variant
Set objProd = CATIA.ActiveDocument.Product.Products

'Create text file in C:\temp\ directory.
Open OutputFile For Output As fnum1

Print #fnum1, "Item Index" & "," & "Part Name" & "," & "Type"
Print #fnum1, "" & "," & "" & "," & objProd.Count

'loop through all parts in feature tree
For i = 1 To objProd.Count
     
    Print #fnum1, i & "," & objProd.Item(i).Name & "," & objProd.Item(i).Parameters.Item("Type").ValueAsString
     
Next
 
'Close the freefile
Close fnum1
 
End Sub
