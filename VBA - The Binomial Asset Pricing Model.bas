Attribute VB_Name = "Module1"
Function BAPM_Amr(S, K, sigma, r, T, q, CP, N)
'All the variables are similar as the Black-Scholes except: N - # of Steps.
'We also redefine dividend as q as we have a D here as down step.
'We are pricing an American Option using the Binomial Asset Pricing Model.

'First we define some variables: U - Up Step, D - Down Step, P - Probability of Occurence.
Dim U, D As Double
ReDim V(N + 1) As Double
Dim fl As Long

U = Exp(sigma * (T / N) ^ 0.5)
D = 1 / U
'The Up Step (U) has a probability of P of occuring & Down Step (D) has a probability of 1 - P of occuring.
P = (Exp((r - q) * T / N) - D) / (U - D)

If CP = "Call" Then
    fl = 1
ElseIf CP = "Put" Then
    fl = -1
End If

'Here we define the Terminal Values of the Binomial Asset Pricing Model.rr
For i = 0 To N 'Taking the example of a 3 Step Tree, we have 4 Terminal Values.
    V(i + 1) = WorksheetFunction.Max(fl * (S * U ^ (N - i) * D ^ i - K), 0)
Next i

'Here we working backwards from the Terminal Values of the Binomial Asset Pricing Model.
'We have the j loop for the number of steps we take backwards.
For j = N To 1 Step -1
    'We have the l loop for the creation of a vector containing the values at that current time.
    For l = 1 To j
        ExpV = (V(l) * P + V(l + 1) * (1 - P)) * Exp(-(r - q) * T / N) 'Expectation Value
        ExrV = WorksheetFunction.Max(fl * (S * U ^ (j - l) * D ^ (l - 1) - K), 0) 'Exercise Value
        V(l) = WorksheetFunction.Max(ExpV, ExrV) 'We take what is higher - Expectation/Exercise (Because we can choose in an American Option).
    Next l
Next j

BAPM_Amr = V(1)

End Function
