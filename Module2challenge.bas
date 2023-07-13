Attribute VB_Name = "Module2"
Sub Summary_Stocksheet():

    'set dimension
    Dim Ticker As String
    Dim Ticker_volume As Double
    Dim Lastrow As Long
    Dim i As Long
    Dim j As Integer
    

    'Headers for Summary sheet
    Range("I1, P1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Define Ticker variable
    Ticker = " "
    Ticker_volume = 0
    

    'Set variables for prices and percent changes
    Dim open_price As Double
    open_price = 0
    Dim close_price As Double
    close_price = 0
    Dim price_change As Double
    price_change = 0
    Dim price_change_percent As Double
    price_change_percent = 0

    
    'For loop
    For i = 2 To Lastrow
    
    'Ticker symbol output
  
    
    
End Sub

