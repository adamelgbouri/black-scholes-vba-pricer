Attribute VB_Name = "Module1"
Option Explicit

Sub RunBlackScholes()

Dim S As Double, K As Double, T As Double
Dim r As Double, sigma As Double, q As Double
    
Dim d1 As Double, d2 As Double
Dim Nd1 As Double, Nd2 As Double
Dim pdf_d1 As Double

Dim CallPrice As Double, PutPrice As Double
Dim DeltaCall As Double, DeltaPut As Double
Dim Gamma As Double, Vega As Double
Dim ThetaCall As Double, ThetaPut As Double
Dim RhoCall As Double, RhoPut As Double
    
' ===== INPUTS =====
With Sheets("BSM")
S = .Range("C4").Value
K = .Range("C5").Value
T = .Range("C6").Value
r = .Range("C7").Value
sigma = .Range("C8").Value
q = .Range("C9").Value
End With
    
' ===== CALCULS =====
d1 = (Log(S / K) + (r - q + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
d2 = d1 - sigma * Sqr(T)
    
Nd1 = Application.WorksheetFunction.Norm_S_Dist(d1, True)
Nd2 = Application.WorksheetFunction.Norm_S_Dist(d2, True)
pdf_d1 = Application.WorksheetFunction.Norm_S_Dist(d1, False)
    
CallPrice = S * Exp(-q * T) * Nd1 - K * Exp(-r * T) * Nd2
    PutPrice = K * Exp(-r * T) * _
               Application.WorksheetFunction.Norm_S_Dist(-d2, True) _
               - S * Exp(-q * T) * _
               Application.WorksheetFunction.Norm_S_Dist(-d1, True)
    
DeltaCall = Exp(-q * T) * Nd1
DeltaPut = DeltaCall - Exp(-q * T)
    
Gamma = Exp(-q * T) * pdf_d1 / (S * sigma * Sqr(T))

Vega = S * Exp(-q * T) * Sqr(T) * pdf_d1
    
    ThetaCall = -(S * Exp(-q * T) * pdf_d1 * sigma) / (2 * Sqr(T)) _
                - r * K * Exp(-r * T) * Nd2 _
                + q * S * Exp(-q * T) * Nd1
    
    ThetaPut = -(S * Exp(-q * T) * pdf_d1 * sigma) / (2 * Sqr(T)) _
               + r * K * Exp(-r * T) * _
               Application.WorksheetFunction.Norm_S_Dist(-d2, True) _
               - q * S * Exp(-q * T) * _
               Application.WorksheetFunction.Norm_S_Dist(-d1, True)
    
    RhoCall = K * T * Exp(-r * T) * Nd2
    RhoPut = -K * T * Exp(-r * T) * _
        Application.WorksheetFunction.Norm_S_Dist(-d2, True)
    
    ' ===== ÉCRITURE RÉSULTATS =====
With Sheets("BSM")
.Range("B15").Value = d1
.Range("B16").Value = d2
.Range("B22").Value = CallPrice
.Range("B23").Value = PutPrice
.Range("D28").Value = DeltaCall
.Range("D29").Value = DeltaPut
.Range("D31").Value = Gamma
.Range("D33").Value = Vega
.Range("D35").Value = ThetaCall
.Range("D36").Value = ThetaPut
.Range("D38").Value = RhoCall
.Range("D39").Value = RhoPut
End With
    
MsgBox "Black-Scholes updated successfully!", vbInformation

End Sub


