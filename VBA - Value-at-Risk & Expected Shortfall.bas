Attribute VB_Name = "Module1"
Function ValAtRisk(obs, alpha, dist, VC)
'obs is observations
'alpha is the confidence level
'dist is the type of distribution (Loss or Profit and Loss)
'VC is choosing VaR or Conditional VaR/Expected Shortfall

If dist = "L" Then
    'L stands for a Loss Distribution
    ValAtRisk = Application.WorksheetFunction.Percentile_Inc(obs, alpha)
    If VC = "VaR" Then
        MsgBox "We have a " & alpha * 100 & "% chance of losing less than " & ValAtRisk & "€ in a day for a Loss distribution."
    ElseIf VC = "CVaR" Then
    'We are modelling for the Expected Shortfall (Loss Distribution)
        Average = Application.WorksheetFunction.SumIf(obs, "<" & ValAtRisk)
        Count = Application.WorksheetFunction.CountIf(obs, "<" & ValAtRisk)
        ValAtRisk = Average / Count
        MsgBox "We have an average of " & ValAtRisk & "€ of losses in the top " & (1 - alpha) * 100 & "% of scenarios for a Loss distribution."
    End If
ElseIf dist = "PnL" Then
    'PnL stands for a Profit and Loss Distribution
    ValAtRisk = Application.WorksheetFunction.Percentile_Inc(obs, 1 - alpha)
    If VC = "VaR" Then
        MsgBox "We have a " & alpha * 100 & "% chance of losing less than " & ValAtRisk & "€ in a day for a PnL distribution."
    ElseIf VC = "CVaR" Then
    'We are modelling for the Expected Shortfall (PnL Distribution)
        Average = Application.WorksheetFunction.SumIf(obs, "<" & ValAtRisk)
        Count = Application.WorksheetFunction.CountIf(obs, "<" & ValAtRisk)
        ValAtRisk = Average / Count
        MsgBox "We have an average of " & ValAtRisk & "€ of losses in the top " & (1 - alpha) * 100 & "% of scenarios for a PnL distribution."
    End If
End If

End Function




    

