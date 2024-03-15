Attribute VB_Name = "Module1"
Sub CalculateAndWriteData()
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentStock As String
    Dim currentYear As Integer
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    ' Create a new worksheet for output
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOutput.Name = "OutputData"
    
    ' Write headers
    wsOutput.Cells(1, 1).Value = "Stock Name"
    wsOutput.Cells(1, 2).Value = "Yearly Change"
    wsOutput.Cells(1, 3).Value = "Percentage Change"
    wsOutput.Cells(1, 4).Value = "Total Stock Volume"
    
    ' Loop through each worksheet
    For Each wsInput In ThisWorkbook.Sheets
        ' Skip the output worksheet and any other sheet that is not needed
        If wsInput.Name = "OutputData" Then
            GoTo ContinueLoop
        End If
        
        ' Find the last row with data in the input worksheet
        lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        currentStock = wsInput.Cells(2, 1).Value ' Assuming stock name is in column A and starts from row 2
        currentYear = Year(DateValue(Left(wsInput.Cells(2, 2).Value, 4) & "/" & Mid(wsInput.Cells(2, 2).Value, 5, 2) & "/" & Right(wsInput.Cells(2, 2).Value, 2))) ' Assuming date is in column B
        openingPrice = wsInput.Cells(2, 3).Value ' Assuming opening price is in column C
        totalVolume = 0
        outputRow = 2 ' Start writing data from row 2
        
        ' Loop through each row and calculate yearly change, percentage change, and total volume
        For i = 2 To lastRow ' Assuming data starts from row 2
            If wsInput.Cells(i, 1).Value <> currentStock Or Year(DateValue(Left(wsInput.Cells(i, 2).Value, 4) & "/" & Mid(wsInput.Cells(i, 2).Value, 5, 2) & "/" & Right(wsInput.Cells(i, 2).Value, 2))) <> currentYear Then
                ' Calculate yearly change and percentage change
                If currentYear <> 0 Then
                    closingPrice = wsInput.Cells(i - 1, 6).Value ' Assuming closing price is in column F
                    yearlyChange = closingPrice - openingPrice
                    If openingPrice <> 0 Then
                        percentageChange = (yearlyChange / openingPrice) * 100
                    Else
                        percentageChange = 0
                    End If
                    
                    ' Write data to output worksheet
                    wsOutput.Cells(outputRow, 1).Value = currentStock
                    wsOutput.Cells(outputRow, 2).Value = yearlyChange
                    wsOutput.Cells(outputRow, 3).Value = percentageChange
                    wsOutput.Cells(outputRow, 4).Value = totalVolume
                    
                    ' Format "Yearly Change" cell based on its value
                    If yearlyChange < 0 Then
                        wsOutput.Cells(outputRow, 2).Interior.Color = RGB(255, 0, 0) ' Red color
                    Else
                        wsOutput.Cells(outputRow, 2).Interior.Color = RGB(0, 255, 0) ' Green color
                    End If
                    
                    ' Format "Percentage Change" cell based on its value
                    If percentageChange < 0 Then
                        wsOutput.Cells(outputRow, 3).Interior.Color = RGB(255, 0, 0) ' Red color
                    Else
                        wsOutput.Cells(outputRow, 3).Interior.Color = RGB(0, 255, 0) ' Green color
                    End If
                    
                    ' Move to the next row in the output worksheet
                    outputRow = outputRow + 1
                End If
                
                ' Update variables for the new stock or year
                currentStock = wsInput.Cells(i, 1).Value
                currentYear = Year(DateValue(Left(wsInput.Cells(i, 2).Value, 4) & "/" & Mid(wsInput.Cells(i, 2).Value, 5, 2) & "/" & Right(wsInput.Cells(i, 2).Value, 2)))
                openingPrice = wsInput.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' Accumulate total volume
            totalVolume = totalVolume + wsInput.Cells(i, 7).Value ' Assuming volume is in column G
        Next i
        
        ' Write data for the last stock
        closingPrice = wsInput.Cells(lastRow, 6).Value
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentageChange = (yearlyChange / openingPrice) * 100
        Else
            percentageChange = 0
        End If
        wsOutput.Cells(outputRow, 1).Value = currentStock
        wsOutput.Cells(outputRow, 2).Value = yearlyChange
        wsOutput.Cells(outputRow, 3).Value = percentageChange
        wsOutput.Cells(outputRow, 4).Value = totalVolume
        
        ' Format "Yearly Change" cell based on its value for the last row
        If yearlyChange < 0 Then
            wsOutput.Cells(outputRow, 2).Interior.Color = RGB(255, 0, 0) ' Red color
        Else
            wsOutput.Cells(outputRow, 2).Interior.Color = RGB(0, 255, 0) ' Green color
        End If
        
        ' Format "Percentage Change" cell based on its value for the last row
        If percentageChange < 0 Then
            wsOutput.Cells(outputRow, 3).Interior.Color = RGB(255, 0, 0) ' Red color
        Else
            wsOutput.Cells(outputRow, 3).Interior.Color = RGB(0, 255, 0) ' Green color
        End If
        
ContinueLoop:
    Next wsInput
End Sub

