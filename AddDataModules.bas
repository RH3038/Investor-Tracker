Attribute VB_Name = "Module3"
Sub AddData()

'Calls each function to update sheets and displays msgBox prompt when completed
Call AddPurchaseData
Call AddCurrentHoldings
Call AddCombinedCurrentHoldings
Call ClearContents1
MsgBox ("Data Has Been Successfully Entered")

End Sub

Sub AddPurchaseData()

'Eliminates the loading time for the application for a seamless update
Application.ScreenUpdating = False

Dim count As Integer
Dim lastRow As Long
Dim writeRow As Long

'Check to see if any entries are blank
count = Application.WorksheetFunction.CountBlank(Sheets("Buy Data Entry").Range("H6:H11"))

If count > 0 Then
MsgBox ("Please fill out entire form to make an entry!")
End

Else

'Inserts a blank row into the table that is then updated with BuyData information entered above the last row
    Sheets("Purchase Data").Select
    lastRow = Sheets("Purchase Data").Cells(Rows.count, 1).End(xlUp).Row
    Rows(lastRow).Select
    Selection.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
    
End If

'Assign data from Data Entry to Purchase Data
lastRow = Sheets("Purchase Data").Cells(Rows.count, 1).End(xlUp).Row

'Finds the empty row that was created and stores it in variable
For l = 1 To lastRow
    If Worksheets("Purchase Data").Cells(l, 1).value = "" Then
    writeRow = l
    End If
Next l

'Sets the entries from Data Entry into Purchase data blank cells

    Worksheets("Buy Data Entry").Range("H6:H11").Copy
    Worksheets("Purchase Data").Range("A" & writeRow & ":" & "F" & writeRow).PasteSpecial Paste:=xlPasteValues, Transpose:=True

Application.ScreenUpdating = True

End Sub

Sub AddCurrentHoldings()

'Eliminates the loading time for the application for a seamless update
Application.ScreenUpdating = False

Dim count As Integer
Dim lastRow As Long
Dim writeRow As Long

'Check to see if any entries are blank
count = Application.WorksheetFunction.CountBlank(Sheets("Buy Data Entry").Range("H6:H11"))

If count > 0 Then
MsgBox ("Please fill out entire form to make an entry!")
End

Else
'Inserts a blank row into the table that is then updated with BuyData information entered above the last row
    Sheets("Current Holdings").Select
    lastRow = Sheets("Current Holdings").Cells(Rows.count, 1).End(xlUp).Row
    Rows(lastRow).Select
    Selection.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
    
End If

'Assign data from Data Entry to Purchase Data
lastRow = Sheets("Current Holdings").Cells(Rows.count, 1).End(xlUp).Row

'Finds the empty row that was created and stores it in variable
For l = 1 To lastRow
    If Worksheets("Current Holdings").Cells(l, 1).value = "" Then
    writeRow = l
    End If
Next l

'Sets the empty cells with data from "Buy Data Entry" H6:H11 to last row in column ranges A:F
    Worksheets("Buy Data Entry").Range("H6:H11").Copy
    Worksheets("Current Holdings").Range("A" & writeRow & ":" & "F" & writeRow).PasteSpecial Paste:=xlPasteValues, Transpose:=True
    
'Pastes appropriate formulas into the last row of columns G:K
    Worksheets("Current Holdings").Range("G2").Copy
    Worksheets("Current Holdings").Range("G" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Worksheets("Current Holdings").Range("H2").Copy
    Worksheets("Current Holdings").Range("H" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Worksheets("Current Holdings").Range("I2").Copy
    Worksheets("Current Holdings").Range("I" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Worksheets("Current Holdings").Range("J2").Copy
    Worksheets("Current Holdings").Range("J" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Worksheets("Current Holdings").Range("K2").Copy
    Worksheets("Current Holdings").Range("K" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False
    
    'Pastes appropriate formulas into the last row of columns G:K and updates the appropriate rolling range
    Worksheets("Current Holdings").Range("H" & lastRow).Formula = "=SUM(" & Range(Cells(2, 8), Cells(lastRow - 1, 8)).Address(False, False) & ")"
    Worksheets("Current Holdings").Range("I" & lastRow).Formula = "=SUM(" & Range(Cells(2, 9), Cells(lastRow - 1, 9)).Address(False, False) & ")" & " / " & lastRow - 2
    Worksheets("Current Holdings").Range("J" & lastRow).Formula = "=SUM(" & Range(Cells(2, 10), Cells(lastRow - 1, 10)).Address(False, False) & ")"
    Worksheets("Current Holdings").Range("K" & lastRow).Formula = "=SUM(" & Range(Cells(2, 11), Cells(lastRow - 1, 11)).Address(False, False) & ")"
    
Application.ScreenUpdating = True

End Sub

Sub AddCombinedCurrentHoldings()

'Eliminates the loading time for the application for a seamless update
Application.ScreenUpdating = False

'Get the last row with investor data and store it in variable
Dim lastRow As Long

Sheets("Combined Current Holdings").Select
lastRow = Sheets("Combined Current Holdings").Cells(Rows.count, 1).End(xlUp).Row

'Variable declared that hold cell data from "Buy Data Entry" sheet within the range H6:H11
Dim buyData() As Variant

'Populate buyData array with text values from "Buy Data Entry" sheet in range H6:H11
For i = 6 To 11
    ReDim Preserve buyData(i - 6)
    buyData(i - 6) = Sheets("Buy Data Entry").Cells(i, 8).Text
Next i

Dim sharesCCH As Variant, avgPriceCCH As Variant, sharesBDE As Variant, priceBDE As Variant, switch As Boolean: switch = False
Dim firLasStock1 As String, firLasStock2 As String

For j = 1 To lastRow

'Stores the First Name, Last Name, and Stock data in variables
    firLasStock1 = buyData(0) & " " & buyData(1) & " " & buyData(2)
    
    firLasStock2 = Trim(Sheets("Combined Current Holdings").Range("A" & j).Text) & " " & _
    Trim(Sheets("Combined Current Holdings").Range("B" & j).Text) & " " & _
    Trim(Sheets("Combined Current Holdings").Range("C" & j).Text)
    
'Statement checks to see if statements are equal and either adds data to existing row if true or _
 creates new row if false
    If firLasStock1 = firLasStock2 Then
        sharesCCH = Sheets("Combined Current Holdings").Range("D" & j).Value2
        avgPriceCCH = Sheets("Combined Current Holdings").Range("F" & j).Value2
        sharesBDE = buyData(3)
        priceBDE = buyData(5)
        purDate = buyData(4)
        switch = True
        
'Adjusts the cells of D?:F? with the appropriate input F: calculates new avg share price _
 D: calcualtes the new amount of shares E: stores the most recent purchase date
         
        Dim totPriceCCH As Variant: totPriceCCH = avgPriceCCH * sharesCCH
        Dim totPriceBDE As Variant: totPriceBDE = priceBDE * sharesBDE
        Dim totShares As Variant: totShares = sharesCCH + sharesBDE
        Dim answer As Variant: answer = (totPriceCCH + totPriceBDE) / totShares
        Sheets("Combined Current Holdings").Range("F" & j) = answer
        
        Sheets("Combined Current Holdings").Range("D" & j) = _
        sharesCCH + sharesBDE
        
        Sheets("Combined Current Holdings").Range("E" & j) = purDate
        
    ElseIf j = lastRow And switch = False Then
    
'Inserts a blank row into the table that is then updated with BuyData information entered above the last row
        Sheets("Combined Current Holdings").Select
        lastRow = Sheets("Combined Current Holdings").Cells(Rows.count, 1).End(xlUp).Row
        Rows(lastRow).Select
        Selection.Insert Shift:=xlUp, CopyOrigin:=FormatFromLeftOrAbove
    
'Selects the range on BuyData worksheet to be updated on other sheets
        Worksheets("Buy Data Entry").Activate
        ActiveSheet.Range("H6:H11").Select
        Selection.Copy
       
'Pastes the above selection into the last row of colums A:F
        Worksheets("Combined Current Holdings").Activate
        ActiveSheet.Range("A" & lastRow & ":" & "F" & lastRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Transpose:=True
       
'Pastes appropriate formulas into the last row of columns G:K
        Worksheets("Combined Current Holdings").Range("G2").Copy
        Worksheets("Combined Current Holdings").Range("G" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
        Worksheets("Combined Current Holdings").Range("H2").Copy
        Worksheets("Combined Current Holdings").Range("H" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
        Worksheets("Combined Current Holdings").Range("I2").Copy
        Worksheets("Combined Current Holdings").Range("I" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
        Worksheets("Combined Current Holdings").Range("J2").Copy
        Worksheets("Combined Current Holdings").Range("J" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
        Worksheets("Combined Current Holdings").Range("K2").Copy
        Worksheets("Combined Current Holdings").Range("K" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        
        Application.CutCopyMode = False
        
'Pastes appropriate formulas into the last row of columns G:K and updates the appropriate rolling range
        Worksheets("Combined Current Holdings").Range("H" & lastRow + 1).Formula = "=SUM(" & Range(Cells(2, 8), Cells(lastRow, 8)).Address(False, False) & ")"
        Worksheets("Combined Current Holdings").Range("I" & lastRow + 1).Formula = "=SUM(" & Range(Cells(2, 9), Cells(lastRow, 9)).Address(False, False) & ")" & " / " & lastRow - 1
        Worksheets("Combined Current Holdings").Range("J" & lastRow + 1).Formula = "=SUM(" & Range(Cells(2, 10), Cells(lastRow, 10)).Address(False, False) & ")"
        Worksheets("Combined Current Holdings").Range("K" & lastRow + 1).Formula = "=SUM(" & Range(Cells(2, 11), Cells(lastRow, 11)).Address(False, False) & ")"
        
    End If
    
Next j

Application.ScreenUpdating = True

End Sub




