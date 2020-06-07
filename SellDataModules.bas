Attribute VB_Name = "Module1"
Sub SubData()

Call AddSellData
Call SubCurrentHoldings
Call SubCombinedCurrentHoldings
Call ClearContents2


End Sub

Sub AddSellData()

'Eliminates the loading time for the application for a seamless update
    Application.ScreenUpdating = False

    Dim count As Integer
    Dim lastRow As Long
    Dim writeRow As Long

'Check to see if any entries are blank
    count = Application.WorksheetFunction.CountBlank(Sheets("Sell Data Entry").Range("H6:H11"))

    If count > 0 Then
        MsgBox ("Please fill out entire form to make an entry!")
    End

    Else

'Inserts a blank row into the table that is then updated with SellDataEntry information entered above the last row
    Sheets("Sell Data").Select
    lastRow = Sheets("Sell Data").Cells(Rows.count, 1).End(xlUp).Row
    Rows(lastRow).Select
    Selection.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
    
    End If

'Assign data from Sell Data Entry to Sell Data
    lastRow = Sheets("Sell Data").Cells(Rows.count, 1).End(xlUp).Row

'Finds the empty row that was created and stores it in variable
    For l = 1 To lastRow
       If Worksheets("Sell Data").Cells(l, 1).value = "" Then
            writeRow = l
       End If
    Next l

'Sets the entries from Data Entry into Purchase data blank cells

    Worksheets("Sell Data Entry").Range("H6:H11").Copy
    Worksheets("Sell Data").Range("A" & writeRow & ":" & "F" & writeRow).PasteSpecial Paste:=xlPasteValues, Transpose:=True
    
'Variable declared that hold cell data from "Buy Data Entry" sheet within the range H6:H11
    Dim sellData() As Variant

'Populate buyData array with text values from "Buy Data Entry" sheet in range H6:H11
    For i = 6 To 11
        ReDim Preserve sellData(i - 6)
        sellData(i - 6) = Sheets("Sell Data Entry").Cells(i, 8).Text
    Next i
    
    Dim firLasStock1 As String, firLasStock2 As String
    Dim switch As Boolean: switch = False

'Finds and updates Sell Data "Purchase Price" with CCH data
    For j = 2 To lastRow

'Stores the First Name, Last Name, and Stock data in variables
        firLasStock1 = sellData(0) & " " & sellData(1) & " " & sellData(2)
    
        firLasStock2 = Sheets("Combined Current Holdings").Range("A" & j).Text & " " & _
        Sheets("Combined Current Holdings").Range("B" & j).Text & " " & _
        Sheets("Combined Current Holdings").Range("C" & j).Text
    
'Statement checks to see if statements are equal and either adds "Purchase Price" to column G in _
Sell Data or retruns error message
        If firLasStock1 = firLasStock2 Then
            Sheets("Combined Current Holdings").Cells(j, 6).Copy
            Sheets("Sell Data").Cells(lastRow - 1, 7).PasteSpecial Paste:=xlPasteValues
            switch = True
        
'If the entry does not exist of CCH then the row that was created is deleted and MsgBox promts error!
        ElseIf j = lastRow And switch = False Then
            Worksheets("Sell Data Entry").Select
            MsgBox "Error: Data entered does not exist"
            Sheets("Sell Data").Rows(lastRow - 1).Delete
            End
        End If
    
    Next j
    
'Pastes appropriate formulas into the last row of columns H:I
    Worksheets("Sell Data").Range("H2").Copy
    Worksheets("Sell Data").Range("H" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Worksheets("Sell Data").Range("I2").Copy
    Worksheets("Sell Data").Range("I" & writeRow).PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True


End Sub

Sub SubCurrentHoldings()

Application.ScreenUpdating = False

'Variables that represent shares for "Current Holdings & Sell Data Entry" as well as string comparisons for those sheets
Dim sharesCH As Variant
Dim sharesSDE As Variant
Dim CH As String
Dim SDE As String

'Selects the "Sell Data Entry" sheet, finds the last row and stores "First name, Last name & Stock" in variable for comparison
SDE = Trim(Sheets("Sell Data Entry").Range("H6").Text) & Trim(Sheets("Sell Data Entry").Range("H7").Text) & _
Trim(Sheets("Sell Data Entry").Range("H8").Text)
sharesSDE = Sheets("Sell Data Entry").Range("H9").Value2

Sheets("Current Holdings").Select

lastRow = Sheets("Current Holdings").Cells(Rows.count, 1).End(xlUp).Row

'Loops through table subtracting share counts and deleting appropritate rows
For i = 2 To lastRow + 1
    CH = Trim(Sheets("Current Holdings").Range("A" & i).Text) & Trim(Sheets("Current Holdings").Range("B" & i).Text) & _
    Trim(Sheets("Current Holdings").Range("C" & i).Text)
    sharesCH = Sheets("Current Holdings").Range("D" & i).Value2
    
    If SDE = CH Then
            
        If sharesSDE = sharesCH Then
            Sheets("Current Holdings").Range("A" & i).EntireRow.Select
            Selection.Delete
            Exit For
        ElseIf sharesSDE > sharesCH Then
            sharesSDE = sharesSDE - sharesCH
            Sheets("Current Holdings").Range("A" & i).EntireRow.Select
            Selection.Delete
            lastRow = Sheets("Current Holdings").Cells(Rows.count, 1).End(xlUp).Row
            i = i - 1
        ElseIf sharesSDE < sharesCH Then
            Sheets("Current Holdings").Cells(i, 4).value = sharesCH - sharesSDE
            Exit For
        End If

    End If

Next i

Application.ScreenUpdating = True

End Sub

Sub SubCombinedCurrentHoldings()

Application.ScreenUpdating = False

'Variables that represent shares for "Combined Current Holdings & Sell Data Entry" as well as string comparisons for those sheets
Dim sharesCCH As Variant
Dim sharesSDE As Variant
Dim CCH As String
Dim SDE As String

'Selects the "Sell Data Entry" sheet, finds the last row and stores "First name, Last name & Stock" in variable for comparison
SDE = Trim(Sheets("Sell Data Entry").Range("H6").Text) & Trim(Sheets("Sell Data Entry").Range("H7").Text) & _
Trim(Sheets("Sell Data Entry").Range("H8").Text)
sharesSDE = Sheets("Sell Data Entry").Range("H9").Value2

Sheets("Combined Current Holdings").Select

lastRow = Sheets("Combined Current Holdings").Cells(Rows.count, 1).End(xlUp).Row

'Loops through table subtracting share counts and deleting appropritate rows
For i = 2 To lastRow + 1
    CCH = Trim(Sheets("Combined Current Holdings").Range("A" & i).Text) & Trim(Sheets("Combined Current Holdings").Range("B" & i).Text) & _
    Trim(Sheets("Combined Current Holdings").Range("C" & i).Text)
    sharesCCH = Sheets("Combined Current Holdings").Range("D" & i).Value2
    
    If SDE = CCH Then
            
        If sharesSDE = sharesCCH Then
            Sheets("Combined Current Holdings").Range("A" & i).EntireRow.Select
            Selection.Delete
            Exit For
        ElseIf sharesSDE > sharesCCH Then
            MsgBox "Error: Share count enetered exceeds held amount"
        ElseIf sharesSDE < sharesCCH Then
            Sheets("Combined Current Holdings").Cells(i, 4).value = sharesCCH - sharesSDE
            Exit For
        End If

    End If

Next i

Application.ScreenUpdating = True

End Sub

