'ENGR 344 - Civil Engineering Analysis
'Winter 2016
'Written by Zachri Jensen

Attribute VB_Name = "modStatistics"
Sub modStatistics_Click()
'Takes stress data from form and determines strain

'Declare Variables
Dim I As Integer
Dim J As Integer
Dim N As Integer
Dim S As Integer

Dim C As Double
Dim MoI As Double
Dim MoE As Double
Dim Length As Double
Dim Pcr(2001) As Double
Dim IMean As Double
Dim IStd As Double
Dim EMean As Double
Dim EStd As Double
Dim LengthMean As Double
Dim LengthStd As Double
Dim Swap As Double
Dim Pcr5 As Double
Dim Pcr95 As Double
Dim N1, N2, N3, U1, U2, U3, U4 As Double

C1 = Atn(1#)
C = 0.5 / Atn(1#)
N = Sheet1.Cells(5, 3).Value
S = Int(N / 100)
LengthMean = Sheet1.Cells(10, 3).Value
LengthStd = Sheet1.Cells(10, 6).Value
IMean = Sheet1.Cells(15, 3).Value
IStd = Sheet1.Cells(15, 6).Value
EMean = Sheet1.Cells(20, 3).Value
EStd = Sheet1.Cells(20, 6).Value


'Initialize random number sequence
Randomize

'Generate a random normal variates
For I = 1 To N
 U1 = Rnd
 U2 = Rnd
 U3 = Rnd
 U4 = Rnd
 N1 = Sqr(-2# * Log(U1)) * Cos(U2 / C)
 N2 = Sqr(-2# * Log(U1)) * Sin(U2 / C)
 N3 = Sqr(-2# * Log(U3)) * Cos(U4 / C)
 
 Length = 2 * (LengthStd * N1 + LengthMean)
 MoI = IStd * N2 + IMean
 MoE = (EStd * N3 + EMean) * 1000000#
 
 Pcr(I) = 4# * MoI * MoE / (Length * Length * C * C)
Next


'Bubble Sort
For I = 1 To N - 1
 For J = 1 To N - I
  If Pcr(J) > Pcr(J + 1) Then
   Swap = Pcr(J)
   Pcr(J) = Pcr(J + 1)
   Pcr(J + 1) = Swap
  End If
 Next
Next

For I = 1 To N
  Sheet3.Cells(I, 1).Value = Pcr(I)
Next
  
Sheet1.Cells(33, 3).Value = Pcr(S * 5 + 1)
Sheet1.Cells(33, 6).Value = Pcr(N - S * 5)

End Sub

