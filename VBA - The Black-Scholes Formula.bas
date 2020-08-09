Attribute VB_Name = "Module1"
Function BSF(S, K, sigma, r, T, d, CP)
'S is the current Stock Price
'K is the Strike Price
'sigma is the Annualized Volatility
'r is the Risk-free Rate
'T is the Time to Maturity (In Years)
'd is the Annualized Dividends
'PC is the Type of Option (Call/Put)


Dim fl As Long

d1 = (Log(S / K) + (r - d + sigma ^ 2 / 2) * T) / (sigma * T ^ 0.5)
d2 = d1 - sigma * T ^ 0.5

'We must put the flag before the formulas containing fl.
If CP = "Call" Then
    fl = 1
ElseIf CP = "Put" Then
    fl = -1
End If

Nd1 = Application.WorksheetFunction.Norm_S_Dist(fl * d1, True)
Nd2 = Application.WorksheetFunction.Norm_S_Dist(fl * d2, True)

BSF = fl * (S * Exp(-d * T) * Nd1 - Nd2 * K * Exp(-r * T))

End Function
