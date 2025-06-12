Attribute VB_Name = "ScaleFactor"
' Topic; Scale Factor
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 26/03/2022
'

'Compute Grid Point Scale Factor
Function PointGSF(semi_a, df, Ei, Ni)
    
'Projection (Thailand Zone47)
    k0 = 0.9996
    E0 = 500000
    N0 = 0
    
'Ellipsoid Constants
    F = 1 / df
    semi_b = semi_a * (1 - F)
    ee1 = Sqr((semi_a ^ 2 - semi_b ^ 2) / semi_a ^ 2)
    ee2 = Sqr((semi_a ^ 2 - semi_b ^ 2) / semi_b ^ 2)

'Meridian ellipe
    m = (Ni - N0) / k0
    u = m / (semi_a * (1 - (ee1 ^ 2 / 4) - (3 * ee1 ^ 4 / 64) - (5 * ee1 ^ 6 / 256)))
    ee3 = (1 - Sqr(1 - ee1 ^ 2)) / (1 + Sqr(1 - ee1 ^ 2))
    Q1 = u + ((3 * ee3 / 2) - (27 * ee3 ^ 3 / 32)) * Sin(2 * u) + ((21 * ee3 ^ 2 / 16) - (55 * ee3 ^ 4 / 32)) * Sin(4 * u) + (151 * ee3 ^ 3 / 96) * Sin(6 * u) + (1097 * ee3 ^ 4 / 512) * Sin(8 * u)
    v = semi_a / Sqr(1 - ee1 ^ 2 * Sin(Q1) ^ 2)
    nn1 = Sqr(ee2 ^ 2 * Cos(Q1) ^ 2)

'Compute Grid Point Scale Factor
    k = k0 * (1 + ((1 + nn1 ^ 2) / 2) * ((Ei - E0) / (k0 * v)) ^ 2 + ((1 + 6 * nn1 ^ 2) / 24) * ((Ei - E0) / (k0 * v)) ^ 4)

    PointGSF = k
    
End Function

'Compute Line Scale Factor by Simpson's Rule
Function LineGSF(k1, k2, km)

    LineGSF = (k1 + 4 * km + k2) / 6

End Function





