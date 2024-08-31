# VBA-challenge

Stock Analysis VBA Script

Overview

This VBA script is designed to perform stock analysis across all worksheets in all open Excel workbooks. The script calculates the following metrics for each stock:

	1.	Ticker Symbol: Extracted from column A.
	2.	Quarterly Change: The difference between the opening price at the beginning of a quarter and the closing price at the end of that quarter.
	3.	Percentage Change: The percentage difference between the opening and closing prices for the quarter.
	4.	Total Stock Volume: The sum of all volumes for the stock during the quarter.

Additionally, the script identifies and outputs the stocks with the Greatest Percentage Increase, Greatest Percentage Decrease, and Greatest Total Volume for each worksheet.

Features

	•	Multi-Workbook and Worksheet Support: The script runs across all worksheets in all open Excel workbooks simultaneously.
	•	Dynamic Data Handling: Automatically processes data from varying lengths and formats across multiple sheets.
	•	Output Organization: Results are outputted starting from the second row in columns I through L on each worksheet, maintaining clear and organized data.

Prerequisites

	•	Microsoft Excel with VBA enabled.
	•	Ensure that the data is organized as follows in each worksheet:
	•	Column A: Ticker symbol
	•	Column B: Date
	•	Column C: Open price
	•	Column F: Close price
	•	Column G: Volume

Instructions for Use

1. Prepare Your Excel Files

Open all the Excel files you want to analyze. Ensure that each worksheet within these files follows the correct data format mentioned above.

2. Access the VBA Editor

	1.	Open Microsoft Excel.
	2.	Press Alt + F11 to open the VBA editor.

3. Insert the VBA Script

	1.	In the VBA editor, click on Insert > Module to create a new module.
	2.	Copy and paste the following VBA script into the module window:

 3.	Sub StockAnalysisAcrossAllWorkbooks()

    ' Define variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    Dim outputRow As Long

    ' Loop through all open workbooks
    For Each wb In Application.Workbooks
        
        ' Loop through all worksheets in the current workbook
        For Each ws In wb.Worksheets
            ws.Activate
            
            ' Initialize variables
            lastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row ' Find the last row in column B
            totalVolume = 0
            greatestIncrease = 0
            greatestDecrease = 0
            greatestVolume = 0
            outputRow = 2 ' Start output from the second row
            
            ' Initialize the first open price and ticker
            openPrice = ws.Cells(2, 3).Value ' Open price in column C
            ticker = ws.Cells(2, 1).Value ' Ticker symbol in column A
            
            ' Loop through all rows
            For i = 2 To lastRow
                ' Check if the ticker changes (new stock)
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    closePrice = ws.Cells(i, 6).Value ' Close price in column F
                    totalVolume = totalVolume + ws.Cells(i, 7).Value ' Volume in column G
                    
                    ' Calculate quarterly change and percentage change
                    quarterlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = (quarterlyChange / openPrice) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Print results in the specified columns starting from row 2
                    ws.Cells(outputRow, 9).Value = ticker ' Output ticker symbol in column I
                    ws.Cells(outputRow, 10).Value = quarterlyChange ' Output quarterly change in column J
                    ws.Cells(outputRow, 11).Value = percentChange ' Output percentage change in column K
                    ws.Cells(outputRow, 12).Value = totalVolume ' Output total volume in column L
                    
                    ' Move to the next output row
                    outputRow = outputRow + 1
                    
                    ' Determine greatest increase, decrease, and volume
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        tickerIncrease = ticker
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        tickerDecrease = ticker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        tickerVolume = ticker
                    End If
                    
                    ' Reset variables for the next ticker
                    openPrice = ws.Cells(i + 1, 3).Value ' New open price for next ticker
                    ticker = ws.Cells(i + 1, 1).Value ' New ticker from column A
                    totalVolume = 0
                Else
                    ' Accumulate volume for the current ticker
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                End If
            Next i
            
            ' Output greatest increase, decrease, and volume after all data
            ws.Cells(outputRow + 2, 14).Value = "Greatest % Increase"
            ws.Cells(outputRow + 2, 15).Value = tickerIncrease
            ws.Cells(outputRow + 2, 16).Value = greatestIncrease
            
            ws.Cells(outputRow + 3, 14).Value = "Greatest % Decrease"
            ws.Cells(outputRow + 3, 15).Value = tickerDecrease
            ws.Cells(outputRow + 3, 16).Value = greatestDecrease
            
            ws.Cells(outputRow + 4, 14).Value = "Greatest Total Volume"
            ws.Cells(outputRow + 4, 15).Value = tickerVolume
            ws.Cells(outputRow + 4, 16).Value = greatestVolume

        Next ws ' Next worksheet in the current workbook

    Next wb ' Next workbook in the application

    MsgBox "Stock analysis completed across all worksheets in all open workbooks!"

End Sub

4. Run the VBA Script

	1.	Press F5 or go to Run > Run Sub/UserForm to execute the script.
	2.	The script will process all worksheets in all open workbooks and output the analysis results starting from row 2 in columns I through L of each worksheet.

Output

	•	Column I: Ticker symbol (retrieved from column A).
	•	Column J: Quarterly change.
	•	Column K: Percentage change.
	•	Column L: Total stock volume.
	•	The greatest percentage increase, decrease, and total volume are displayed below the data for each worksheet.

Conclusion

By using this VBA script, you can efficiently analyze stock data across multiple worksheets and workbooks simultaneously in Excel. This tool automates the process, saving time and reducing manual effort.

This README file should provide clear guidance on the purpose, usage, and functionality of the VBA script for anyone who wants to perform stock analysis in Excel.
