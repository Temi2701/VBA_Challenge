Attribute VB_Name = "Module1"
Sub MultipleYearStockData()

 ' Loop through all sheets
    For Each ws In Worksheets

  ' Set an initial variable for the Ticker
  Dim Ticker As String

  ' Set an initial variable for the Yearly Change
  Dim Yearly_Change As Double
 Yearly_Change = 0
  
  ' Set an initial variable for the Opening Price
  Dim Opening_Price As Double
  Opening_Price = 0
  
   ' Set an initial variable for the Closing Price
  Dim Closing_Price As Double
  Closing_Price = 0
  
  ' Set an initial variable for the Percentage Change
  Dim Percentage_Change As Double
  Percentage_Change = 0
  
  ' Set an initial variable for the Total Stock Volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Keep track of the location for each item
  Dim Results_Area As Integer
  Results_Area = 2
  
 ' Count the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all entries
  For i = 2 To lastrow

    ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker_Symbol = Cells(i, 1).Value
      
      'Find Closing Price at end of year
      Closing_Price = Closing_Price + Cells(i, 6).Value
      
      'Find Opening Price
      Opening_Price = Opening_Price + Cells(i, 3).Value

      ' Find the Yearly Change
      Yearly_Change = Closing_Price - Opening_Price
      
      'Find the Percentage Change
      Percentage_Change = (Yearly_Change / Closing_Price) * 100
      
      'Find the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print the Ticker in the Results area
      Range("J" & Results_Area).Value = Ticker_Symbol

      ' Print the Yearly Change to the Results area
      Range("K" & Results_Area).Value = Yearly_Change
      
      ' Print the Percentage Change to the Results area
      Range("L" & Results_Area).Value = Percentage_Change
      
       ' Print the Total Stock Volume to the Results area
       Range("M" & Results_Area).Value = Total_Stock_Volume

      ' Add one to the Results Area
      Results_Area = Results_Area + 1
      
      ' Reset Opening Price
      Opening_Price = 0
      
      ' Reset Closing Price
      Closing_Price = 0
      
      ' Reset the Yearly Change
      Yearly_Change = 0
      
      ' Reset the Percentage Change
      Percentage_Change = 0
      
      ' Reset the Total Stock Volume
      Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

       'Find Closing Price
      Closing_Price = Closing_Price + Cells(i, 6).Value
      
      'Find Opening Price
      Opening_Price = Opening_Price + Cells(i, 3).Value

      ' Find the Yearly Change
      Yearly_Change = Closing_Price - Opening_Price
      
      'Find the Percentage Change
      Percentage_Change = (Yearly_Change / Closing_Price) * 100
      
      'Find the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
      ' Print the Ticker in the Results area
      Range("J" & Results_Area).Value = Ticker_Symbol

      ' Print the Yearly Change to the Results area
      Range("K" & Results_Area).Value = Yearly_Change
      
      ' Print the Percentage Change to the Results area
      Range("L" & Results_Area).Value = Percentage_Change
      
       ' Print the Total Stock Volume to the Results area
       Range("M" & Results_Area).Value = Total_Stock_Volume


    End If

  Next i
  
  MsgBox "Done !"
  
  Dim rng As Range
  Dim Greatest_Percentage_Increase As Double
  
  Set rng = Range("L2:L3001")
  Greatest_Percentage_Increase = WorksheetFunction.Max(rng)
    Range("P4").Value = Greatest_Percentage_Increase
    
Dim Greatest_Percentage_Decrease As Double
Greatest_Percentage_Decrease = WorksheetFunction.Min(rng)
Range("P5").Value = Greatest_Percentage_Decrease

Dim rng_2 As Range
Dim Greatest_Total_Volume As Double

Set rng_2 = Range("M2:M3001")
Greatest_Total_Volume = WorksheetFunction.Max(rng_2)
    Range("P6").Value = Greatest_Total_Volume
  
  Next ws

End Sub


