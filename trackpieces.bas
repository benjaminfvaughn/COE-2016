Attribute VB_Name = "trackpieces"
Public Sub pieces()

Dim currentuser As String
Dim bk As Workbook

'Save current username from system environment variables
currentuser = Environ("username")

'Silence error messages
On Error Resume Next
CATIA.DisplayFileAlerts = False
CATIA.RefreshDisplay = False

'Instantiate/initialize objects and variables.
Set CATIA = GetObject(, "CATIA.Application")
Dim objPart As Variant
Set objPart = CATIA.ActiveDocument.Product

Dim objSel As Variant
Set objSel = CATIA.ActiveDocument.Selection

Dim objPartCollection As Variant
Set objPartCollection = objPart.Products

Dim objSubPartCollection As Variant

CATIA.RefreshDisplay = False

'Create Excel output
Set objEXCELapp = CreateObject("EXCEL.Application")
Set bk = Application.ActiveWorkbook

Worksheets(1).Name = "Track Pieces"

Set sh = bk.Sheets("Track Pieces")

'Clear contents of all cells on the worksheet called "Track Pieces"
sh.Cells.Select
Selection.ClearContents
sh.Range("A1").Select

'Search CATIA feature tree and count the number of parts with a type parameter.
objSel.Search ("Name=Type*,all")
numero = objSel.Count
    
'finds the last cell based on the first row (can change the row by changing "A1" and Cells(1,1) in lines
If IsEmpty(sh.Range("A1").Value) = False Then
    'Checks to see if the first value is empty, otherwise you return something all the way to the right
    'if it isn't empty, then it runs this
    firstempty = sh.Cells(1, 1).End(xlToRight).Column + 1
Else
    'if it is empty, it sets the value to 1
    firstempty = 1
End If

'Initialize variables that will contain the location of part information
labelcol = firstempty
typecol = labelcol + 1
materialcol = labelcol + 2
densitycol = labelcol + 3
volumecol = labelcol + 4

'Sets the header information for each column
sh.Cells(1, labelcol) = "Part Name"
sh.Cells(1, typecol) = "Type"
sh.Cells(1, materialcol) = "Material"
sh.Cells(1, densitycol) = "Density"
sh.Cells(1, volumecol) = "Volume"

'returns the number of parts in the product
sh.Cells(2, typecol) = numero

'cycle through each capable and returns the name and part type
For I = 1 To numero
    sh.Cells(2 + I, labelcol) = objPart.Products.Item(I).Name
    sh.Cells(2 + I, typecol) = objPart.Products.Item(I).Parameters.Item("Type").ValueAsString
    sh.Cells(2 + I, materialcol) = objPart.Products.Item(I).Parameters.Item("Material").ValueAsString
    sh.Cells(2 + I, densitycol) = objPart.Products.Item(I).Parameters.Item("Density").ValueAsString
    sh.Cells(2 + I, volumecol) = objPart.Products.Item(I).Parameters.Item("Volume").ValueAsString
Next I

'saves the workbook and tells excel that the workbook is saved to prevent error messages
bk.Save
bk.Saved = True

'objEXCELApp.DisplayAlerts = False
objEXCELapp.Workbooks.Close

CATIA.RefreshDisplay = True
CATIA.DisplayFileAlerts = True

End Sub

