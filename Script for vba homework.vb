Sub Worksheet_Loop()
 
    ' ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
        For Each ws In Worksheets
        
' ' --------------------------------------------
        ' EXTRACT THE WORKSHEET NAME
        ' --------------------------------------------
        ' Created a Variable to Hold File Name, Last Row and Last Column
        Dim WorksheetName As String
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
                  
           ' Add the word Ticker to the H Column Header
        ws.Cells(1, 8).Value = "Ticker"
                  
           ' Add the words Total Stock Volume to the I Column Header
        ws.Cells(1, 9).Value = "Total Stock Volume"
           
        ' Set an initial variable for holding the ticker name
          Dim Ticker_Name As String
          ' Set an initial variable for holding the total per stock ticker name
          Dim Stock_Total As Double
          Stock_Total = 0
          ' Keep track of the location for each stock ticker name in the summary table
          Dim Summary_Table_Row As Integer
          Summary_Table_Row = 2
          ' Loop through all stock ticker names
          For I = 2 To LastRow
            ' Check if we are still within the same stock ticker name, if it is not...
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
              ' Set the stock ticker name
                 Ticker_Name = Cells(I, 1).Value
              ' Add to the Stock Total
              Stock_Total = Stock_Total + Cells(I, 7).Value
              ' Print the Stock Ticker Name in the Summary Table
              ws.Range("H" & Summary_Table_Row).Value = Ticker_Name
              ' Print the Stock Amount to the Summary Table
              ws.Range("I" & Summary_Table_Row).Value = Stock_Total
              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              ' Reset the Stock Total
              Stock_Total = 0
            ' If the cell immediately following a row is the same ticker..
            Else
              ' Add to the Stock Total
              Stock_Total = Stock_Total + Cells(I, 7).Value
            End If
          Next I
                          
        Next ws

End Sub
