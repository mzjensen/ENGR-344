'ENGR 344 - Civil Engineering Analysis
'Winter 2016
'Written by Zachri Jensen

Attribute VB_Name = "Module1"
Sub Calculate_SN()
'Calculates the structural numbers for a flexible pavement structure

'Declare variable types and names
Dim Pt As Single
Dim R As Integer
Dim S As Integer
Dim SN_new As Double
Dim SN_old As Double
Dim epsilon As Single
Dim ESAL(52) As Double
Dim SN(52) As Double

'Get user inputs for Pt, R, S, minESAL, maxESAL, and targetESAL
Pt = Sheet1.Cells(33, 4).Value
R = Sheet1.Cells(35, 4).Value
S = Sheet1.Cells(37, 4).Value
minESAL = Sheet1.Cells(33, 9).Value
maxESAL = Sheet1.Cells(35, 9).Value
targetESAL = Sheet1.Cells(37, 9).Value

'Check to see if user inputs are within range
If Pt <> 2 And Pt <> 2.5 Then
    MsgBox "The Terminal Serviceability Index should be either 2.0 or 2.5." & vbNewLine & "Please change the input and try again."
    Exit Sub
End If

If R < 1 Or R > 4 Then
    MsgBox "The Regional Factor should be between 1 and 4." & vbNewLine & "Please change the input and try again."
    Exit Sub
End If

If S < 1 Or S > 10 Then
    MsgBox "The Soil Support value should be between 1 and 10." & vbNewLine & "Please change the input and try again."
    Exit Sub
End If

'This check is only here to prevent an error, since later calculations involve the log() function
If minESAL <= 0 Then
    MsgBox "The minimum ESAL value should be greater than 0." & vbNewLine & "Please change the input and try again."
    Exit Sub
End If

'A highway shouldn't experience more than 100 million ESALs during its lifetime
If maxESAL > 100000000 Then
    MsgBox "The minimum ESAL value should be less than 100,000,000." & vbNewLine & "Please change the input and try again."
    Exit Sub
End If

If targetESAL < minESAL Or targetESAL > maxESAL Then
    MsgBox "The target ESAL value should be between the minimum and maximum ESAL values." & vbNewLine & "Please change the input and try again."
    Exit Sub
End If

'Initialize variables
lmin = Log(minESAL)
lmax = Log(maxESAL)

'Calculate structural numbers with corresponding ESALs
For I = 1 To 21
    ESAL(I) = Exp(lmin + (lmax - lmin) * 0.05 * (I - 1))
        
        epsilon = 0.001
        SN_new = 4
        SN_old = 0
        
        Do While Abs(SN_old - SN_new) > epsilon
            SN_old = SN_new
            SN_new = 10 ^ ((Log(ESAL(I)) / Log(10) + 0.2 - (Log((4.2 - Pt) / (4.2 - 1.5)) / Log(10)) / _
                     (0.4 + (1094 / (SN_old + 1) ^ 5.19)) + Log(R) / Log(10) - 0.372 * (S - 3)) / 9.36) - 1
        Loop
    
    SN(I) = SN_new
    Sheet2.Cells(I + 3, 1).Value = ESAL(I)
    Sheet2.Cells(I + 3, 2).Value = SN(I)
Next

'Calculate structural number for target ESAL
epsilon = 0.001
SN_new = 4
SN_old = 0
Do While Abs(SN_old - SN_new) > epsilon
    SN_old = SN_new
    SN_new = 10 ^ ((Log(targetESAL) / Log(10) + 0.2 - (Log((4.2 - Pt) / (4.2 - 1.5)) / Log(10)) / _
             (0.4 + (1094 / (SN_old + 1) ^ 5.19)) + Log(R) / Log(10) - 0.372 * (S - 3)) / 9.36) - 1
Loop

Sheet1.Cells(44, 5).Value = SN_new
Sheet1.Cells(44, 3).Value = targetESAL

End Sub




