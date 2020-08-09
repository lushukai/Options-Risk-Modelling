Attribute VB_Name = "Module1"
Function BSF(S, K, sigma, r, T, d, CP)

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

Function ImpVol(S, K, Price, r, T, d, CP)
'We want to get Implied Volatility from the Market Price of the Option.

'Here we are setting up the Iterative Algorithm to solve for ImpVol because there is no Analytical Solution.
sigma1 = 0.3
BSF1 = BSF(S, K, sigma1, r, T, d, CP)

If BSF1 > Price Then
    sigma2 = sigma1 - 0.001
Else
    sigma2 = sigma1 + 0.001
End If
BSF2 = BSF(S, K, sigma2, r, T, d, CP)
'So far we have defined four new variables - sigma1, BSF1, sigma2, BSF2

'At this point it would be good to draw out the graph of Call Price vs Volatility to Visualize the Iterative Algorithm.
'Contrary to what you see, Call Price vs Volatility is not a Straight Line - That's why I plotted Change in Call Price vs Volatility.
Do
    sigma3 = sigma1 + (Price - BSF1) * ((sigma2 - sigma1) / (BSF2 - BSF1))
    BSF3 = BSF(S, K, sigma3, r, T, d, CP)
    If Abs(BSF1 - Price) > Abs(BSF2 - Price) Then
        sigma1 = sigma3
    Else
        sigma2 = sigma3
    End If
Loop Until Abs(BSF3 - Price) < 10 ^ -1

ImpVol = sigma3

End Function
