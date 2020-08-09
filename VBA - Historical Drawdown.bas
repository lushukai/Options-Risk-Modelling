Attribute VB_Name = "Module1"
Function HDD(p)
'P is the set of historical prices (Use Adj Close).
'Using Adj Close is a better reflection of the Stock price.
'S is the Start and E is the End Date.

'We count the number of historical prices as n.
n = Application.WorksheetFunction.Count(p)
HDD = 0
Dim S As Date
Dim E As Date

'Here we define a double loop to calculate drawdowns in a forward fashion.
For i = 1 To n
    For j = i + 1 To n
        Drop = (p(j) - p(i)) / p(i)
        '"Drop" here is set of historical drawdowns.
        If Drop < HDD Then
            HDD = Drop
        End If
    Next j
Next i

HDD = FormatPercent(HDD, 2)
MsgBox "The Historical Drawdown has been " & HDD & " throughout the price range."

End Function
