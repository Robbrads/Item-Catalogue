'Make an order form from catalogue
Sub FillOrderForm()
Dim inputData As Variant
inputData = InputBox("Choose item row:")
Dim units As Variant
units = InputBox("How many units?")
Dim usernm As String
usernm = Application.UserName
Worksheets("Order Form").Range("B2").Value = units
Worksheets("Order Form").Range("B3").Value = Worksheets("Catalogue").Range("D" & inputData).Value
Worksheets("Order Form").Range("B4").Value = Worksheets("Catalogue").Range("A" & inputData).Value
Worksheets("Order Form").Range("B5").Value = Worksheets("Catalogue").Range("C" & inputData).Value
Worksheets("Order Form").Range("B6").Value = Worksheets("Catalogue").Range("B" & inputData).Value
Worksheets("Order Form").Range("B7").Value = Worksheets("Catalogue").Range("M" & inputData).Value
Worksheets("Order Form").Range("B8").Value = Worksheets("Catalogue").Range("I" & inputData).Value
Worksheets("Order Form").Range("B9").Value = Worksheets("Catalogue").Range("J" & inputData).Value
Worksheets("Order Form").Range("B10").Value = Worksheets("Catalogue").Range("G" & inputData).Value
Worksheets("Order Form").Range("B11").Value = Worksheets("Catalogue").Range("H" & inputData).Value
Worksheets("Order Form").Range("B12").Value = usernm
Worksheets("Order Form").Range("B13").Value = Date

End Sub


'Make a multi-order form from catalogue
Sub FillOrderFormMulti()
Dim numberitems As Variant
numberitems = InputBox("How many items to order?")

If numberitems > 8 Then
    MsgBox ("Maximum number of items is 8")
    Exit sub
End If

'item 1
Dim inputData1 As Variant
inputData1 = InputBox("Choose row for item 1:")
Dim units1 As Variant
units1 = InputBox("How many units of item 1?")
Worksheets("Order Form - Multi").Range("B2").Value = units1
Worksheets("Order Form - Multi").Range("B3").Value = Worksheets("Catalogue").Range("D" & inputData1).Value
Worksheets("Order Form - Multi").Range("B4").Value = Worksheets("Catalogue").Range("A" & inputData1).Value
Worksheets("Order Form - Multi").Range("B5").Value = Worksheets("Catalogue").Range("C" & inputData1).Value
Worksheets("Order Form - Multi").Range("B6").Value = Worksheets("Catalogue").Range("I" & inputData1).Value


'item 2
Dim inputData2 As Variant
inputData2 = InputBox("Choose row for item 2:")
Dim units2 As Variant
units2 = InputBox("How many units of item 2?")
Worksheets("Order Form - Multi").Range("B7").Value = units2
Worksheets("Order Form - Multi").Range("B8").Value = Worksheets("Catalogue").Range("D" & inputData2).Value
Worksheets("Order Form - Multi").Range("B9").Value = Worksheets("Catalogue").Range("A" & inputData2).Value
Worksheets("Order Form - Multi").Range("B10").Value = Worksheets("Catalogue").Range("C" & inputData2).Value
Worksheets("Order Form - Multi").Range("B11").Value = Worksheets("Catalogue").Range("I" & inputData2).Value

'item 3
If numberitems > 2 Then
    Dim inputData3 As Variant
    inputData3 = InputBox("Choose row for item 3:")
    Dim units3 As Variant
    units3 = InputBox("How many units of item 3?")
	Worksheets("Order Form - Multi").Range("B12").Value = units3
	Worksheets("Order Form - Multi").Range("B13").Value = Worksheets("Catalogue").Range("D" & inputData3).Value
	Worksheets("Order Form - Multi").Range("B14").Value = Worksheets("Catalogue").Range("A" & inputData3).Value
	Worksheets("Order Form - Multi").Range("B15").Value = Worksheets("Catalogue").Range("C" & inputData3).Value
	Worksheets("Order Form - Multi").Range("B16").Value = Worksheets("Catalogue").Range("I" & inputData3).Value
End If

'item 4
If numberitems > 3 Then
    Dim inputData4 As Variant
    inputData4 = InputBox("Choose row for item 4:")
    Dim units4 As Variant
    units4 = InputBox("How many units of item 4?")
	Worksheets("Order Form - Multi").Range("B17").Value = units4
	Worksheets("Order Form - Multi").Range("B18").Value = Worksheets("Catalogue").Range("D" & inputData4).Value
	Worksheets("Order Form - Multi").Range("B19").Value = Worksheets("Catalogue").Range("A" & inputData4).Value
	Worksheets("Order Form - Multi").Range("B20").Value = Worksheets("Catalogue").Range("C" & inputData4).Value
	Worksheets("Order Form - Multi").Range("B21").Value = Worksheets("Catalogue").Range("I" & inputData4).Value
End If

'item 5
If numberitems > 4 Then
    Dim inputData5 As Variant
    inputData5 = InputBox("Choose row for item 5:")
    Dim units5 As Variant
    units5 = InputBox("How many units of item 5?")
	Worksheets("Order Form - Multi").Range("B22").Value = units5
	Worksheets("Order Form - Multi").Range("B23").Value = Worksheets("Catalogue").Range("D" & inputData5).Value
	Worksheets("Order Form - Multi").Range("B24").Value = Worksheets("Catalogue").Range("A" & inputData5).Value
	Worksheets("Order Form - Multi").Range("B25").Value = Worksheets("Catalogue").Range("C" & inputData5).Value
	Worksheets("Order Form - Multi").Range("B26").Value = Worksheets("Catalogue").Range("I" & inputData5).Value
End If

'item 6
If numberitems > 5 Then
    Dim inputData6 As Variant
    inputData6 = InputBox("Choose row for item 6:")
    Dim units6 As Variant
    units6 = InputBox("How many units of item 6?")
	Worksheets("Order Form - Multi").Range("B27").Value = units6
	Worksheets("Order Form - Multi").Range("B28").Value = Worksheets("Catalogue").Range("D" & inputData6).Value
	Worksheets("Order Form - Multi").Range("B29").Value = Worksheets("Catalogue").Range("A" & inputData6).Value
	Worksheets("Order Form - Multi").Range("B30").Value = Worksheets("Catalogue").Range("C" & inputData6).Value
	Worksheets("Order Form - Multi").Range("B31").Value = Worksheets("Catalogue").Range("I" & inputData6).Value
End If

'item 7
If numberitems > 6 Then
    Dim inputData7 As Variant
    inputData7 = InputBox("Choose row for item 7:")
    Dim units7 As Variant
    units7 = InputBox("How many units of item 7?")
	Worksheets("Order Form - Multi").Range("B32").Value = units7
	Worksheets("Order Form - Multi").Range("B33").Value = Worksheets("Catalogue").Range("D" & inputData7).Value
	Worksheets("Order Form - Multi").Range("B34").Value = Worksheets("Catalogue").Range("A" & inputData7).Value
	Worksheets("Order Form - Multi").Range("B35").Value = Worksheets("Catalogue").Range("C" & inputData7).Value
	Worksheets("Order Form - Multi").Range("B36").Value = Worksheets("Catalogue").Range("I" & inputData7).Value
End If

'item 8
If numberitems > 7 Then
    Dim inputData8 As Variant
    inputData8 = InputBox("Choose row for item 8:")
    Dim units8 As Variant
    units8 = InputBox("How many units of item 8?")
	Worksheets("Order Form - Multi").Range("B37").Value = units7
	Worksheets("Order Form - Multi").Range("B38").Value = Worksheets("Catalogue").Range("D" & inputData8).Value
	Worksheets("Order Form - Multi").Range("B39").Value = Worksheets("Catalogue").Range("A" & inputData8).Value
	Worksheets("Order Form - Multi").Range("B40").Value = Worksheets("Catalogue").Range("C" & inputData8).Value
	Worksheets("Order Form - Multi").Range("B41").Value = Worksheets("Catalogue").Range("I" & inputData8).Value
End If

Worksheets("Order Form - Multi").Range("B42").Value = Worksheets("Catalogue").Range("B" & inputData1).Value
Worksheets("Order Form - Multi").Range("B43").Value = Worksheets("Catalogue").Range("M" & inputData1).Value
Worksheets("Order Form - Multi").Range("B44").Value = Worksheets("Catalogue").Range("J" & inputData1).Value
Worksheets("Order Form - Multi").Range("B45").Value = Worksheets("Catalogue").Range("G" & inputData1).Value
Worksheets("Order Form - Multi").Range("B46").Value = Worksheets("Catalogue").Range("H" & inputData1).Value
Dim usernm As String
usernm = Application.UserName
Worksheets("Order Form - Multi").Range("B47").Value = usernm
Worksheets("Order Form - Multi").Range("B48").Value = Date
End Sub



'Save order form for email
Sub SaveOrder()
XPath = Application.ActiveWorkbook.Path
Dim nm As String
Dim mth As String
Dim yr As String
yr = Year(Date)
mth = MonthName(Month(Now), True)
nm = Worksheets("Order Form").Range("B4")
ThisWorkbook.Sheets("Order Form").Copy
Application.ActiveWorkbook.SaveAs Filename:=XPath & "\" & nm & " order " & mth & " " & yr & ".xlsx"
ActiveWorkbook.Close
End Sub


'Save multi order form for email
Sub SaveOrderMulti()
XPath = Application.ActiveWorkbook.Path
Dim nm As String
Dim mth As String
Dim yr As String
yr = Year(Date)
mth = MonthName(Month(Now), True)
nm = Worksheets("Order Form - Multi").Range("B42")
ThisWorkbook.Sheets("Order Form - Multi").Copy
Application.ActiveWorkbook.SaveAs Filename:=XPath & "\" & nm & " order " & mth & " " & yr & ".xlsx"
ActiveWorkbook.Close
End Sub


'Clear data in the order form after saving a copy
Sub ClearOrderForm()
Worksheets("Order Form").Range("B2:B13").Value = ""
End Sub

'Clear data in multi order form after saving a copy
Sub ClearOrderFormMulti()
Worksheets("Order Form - Multi").Range("B2:B48").Value = ""
End Sub




'create inventory item on the next row and assign a stock number
Sub CreateInventoryItem()
Dim nextrow As Variant
nextrow = Worksheets("Inventory").Range("A" & Rows.Count).End(xlUp).Row + 1
Dim inputData As Variant
inputData = InputBox("Choose item row:")
Dim dtrec as Date
dtrec = InputBox("Enter date received:")
Dim lot As Variant
lot = InputBox("Manufacturer's lot/batch number:")
Dim units As Variant
units = InputBox("How many units?")
Dim usernm As String
usernm = Application.UserName
Dim Expdt as Date
Expdt = InputBox("Enter expiry date:")
Dim kitinsert as Variant
kitinsert = InputBox("Kit insert received? Y/N")
Dim inspection as Variant
inspection = InputBox("Any damage on delivery? Y/N")

Worksheets("Inventory").Range("A" & nextrow).Value = nextrow - 2
Worksheets("Inventory").Range("B" & nextrow).Value = Worksheets("Catalogue").Range("A" & inputData).Value
Worksheets("Inventory").Range("C" & nextrow).Value = lot
Worksheets("Inventory").Range("D" & nextrow).Value = Worksheets("Catalogue").Range("B" & inputData).Value
Worksheets("Inventory").Range("E" & nextrow).Value = Worksheets("Catalogue").Range("C" & inputData).Value
Worksheets("Inventory").Range("F" & nextrow).Value = Worksheets("Catalogue").Range("D" & inputData).Value
Worksheets("Inventory").Range("G" & nextrow).Value = dtrec
Worksheets("Inventory").Range("H" & nextrow).Value = units
Worksheets("Inventory").Range("I" & nextrow).Value = Expdt
Worksheets("Inventory").Range("L" & nextrow).Value = Kitinsert
Worksheets("Inventory").Range("M" & nextrow).Value = usernm

If inspection ="n" then Worksheets("Inventory").Range("J" & nextrow).Value = "OK"
If inspection ="n" then Worksheets("Inventory").Range("K" & nextrow).Value = "N/A"
If inspection ="y" then 
	Dim damage as String
	Damage = InputBox("Describe damage:")
	Worksheets("Inventory").Range("K" & nextrow).Value = Damage
	Worksheets("Inventory").Range("J" & nextrow).Value = "Damaged"
End if

Worksheets("Inventory").select
End Sub


'Fill out verification form from inventory item
Sub VerifyStock()
Dim inputData As Variant
inputData = InputBox("Choose item row:")
Dim usernm As String
Worksheets("Verification").Range("E2").Value = Worksheets("Inventory").Range("B" & inputData).Value
Worksheets("Verification").Range("E3").Value = Worksheets("Inventory").Range("A" & inputData).Value
Worksheets("Verification").Range("E4").Value = Worksheets("Inventory").Range("C" & inputData).Value
Worksheets("Verification").Range("E5").Value = Worksheets("Inventory").Range("I" & inputData).Value
Worksheets("Verification").Range("E6").Value = Worksheets("Inventory").Range("G" & inputData).Value
Worksheets("Verification").Range("E7").Value = Worksheets("Inventory").Range("H" & inputData).Value
If Worksheets("Inventory").Range("J" & inputData).Value = "OK" Then
    Worksheets("Verification").CheckBoxes("Check Box 1").Value = True
End If
If Worksheets("Inventory").Range("J" & inputData).Value = "Damaged" Then
    Worksheets("Verification").CheckBoxes("Check Box 2").Value = True
End If

If Worksheets("Inventory").Range("J" & inputData).Value = "Damaged" Then
    Worksheets("Verification").Range("E9").Value = Worksheets("Inventory").Range("K" & inputData).Value
End If

If Worksheets("Inventory").Range("L" & inputData).Value = "y" Then
    Worksheets("Verification").CheckBoxes("Check Box 3").Value = True
End If

If Worksheets("Inventory").Range("L" & inputData).Value = "n" Then
    Worksheets("Verification").CheckBoxes("Check Box 4").Value = True
End If

Worksheets("Verification").Select
End Sub



'Clear verification form after use
Sub ClearVerificationForm()
Worksheets("Verification").Range("E2:E40").Value = ""
Worksheets("Verification").CheckBoxes("Check Box 1").Value = False
Worksheets("Verification").CheckBoxes("Check Box 2").Value = False
Worksheets("Verification").CheckBoxes("Check Box 3").Value = False
Worksheets("Verification").CheckBoxes("Check Box 4").Value = False
Worksheets("Verification").CheckBoxes("Check Box 6").Value = False
Worksheets("Verification").CheckBoxes("Check Box 7").Value = False
Worksheets("Verification").CheckBoxes("Check Box 8").Value = False
Worksheets("Verification").CheckBoxes("Check Box 11").Value = False
Worksheets("Verification").CheckBoxes("Check Box 13").Value = False
Worksheets("Verification").CheckBoxes("Check Box 14").Value = False

End Sub





