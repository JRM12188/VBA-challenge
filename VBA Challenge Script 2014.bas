Attribute VB_Name = "Module2"
Sub Stockloup()

    Dim ticker As String
    
    
    Dim ticker_Total As Double
        ticker_Total = 0
        
    Dim Ticker_Percent As Double
        Ticker_Percent = 0
    
    Dim Vol As Double
        Vol = 0
    
    Dim LastRow As Long

        Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

      
 For i = 2 To 705715

    If Cells(i + 1, 1).Value = Cells(i, 1).Value And Cells(i, 2).Value = 20140101 Then
    
    BegValue = Cells(i, 3).Value
    
    Else
    
    
    
    End If

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value
    
    If Cells(i, 2).Value = 20141231 Then
      
      Endvalue = Cells(i, 6).Value
      
      End If

      ticker_Total = ticker_Total + (Endvalue - BegValue)
      
      Ticker_Percent = Ticker_Percent - (1 - (Endvalue / BegValue))
      
      Vol = Vol + Cells(i, 7).Value
      
      Range("I" & Summary_Table_Row).Value = ticker

      Range("J" & Summary_Table_Row).Value = ticker_Total
      
      Range("k" & Summary_Table_Row).Value = Ticker_Percent
      
      Range("L" & Summary_Table_Row).Value = Vol
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      ticker_Total = 0
      
      Ticker_Percent = 0
      
      Vol = 0
  
    
    Else
      
      Vol = Vol + Cells(i, 7).Value
  End If
  
      
Next i


End Sub

