Attribute VB_Name = "Module1"
Function MCAsian(S, K, sigma, r, T, q, CP, N, Nsim)
'Here the variables are the same as BAPM except with the addition of: Nsim, Number of Simulations.

Dim Z, V As Double

If CP = "Call" Then
    fl = 1
ElseIf CP = "Put" Then
    fl = -1
End If

For i = 1 To Nsim 'We are running the Monte Carlo Simulation here.
    St = S
    Av = S
    For j = 1 To N 'We are creating a Stochastic Evolution of Stock Price.
        Randomize
        Z = WorksheetFunction.Norm_S_Inv(Rnd())
        St = St * Exp((r - sigma ^ 2 / 2) * (T / N) + Z * sigma * (T / N) ^ 0.5) 'We use the Black-Scholes Framework for the Stochastic Evolution.
        Av = Av + St
    Next j
    V = V + WorksheetFunction.Max(fl * (Av / (N + 1) - K), 0) 'Here we do the Averaging in an Asian Option to create the payoff.
Next i

MCAsian = V / Nsim * Exp(-r * T) 'Here we take the Average of the Simulations to Calculate Asian Option Price.
'Do note we can increase Nsim, Number of Simulations, for Higher Accuracy in pricing.
'We can also increase N, The Stochastic Evolution, for more randomness.

End Function

