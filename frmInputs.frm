'ENGR 344 - Civil Engineering Analysis
'Winter 2016
'Written by Zachri Jensen

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputs 
   Caption         =   "UserForm1"
   ClientHeight    =   4200
   ClientLeft      =   40
   ClientTop       =   340
   ClientWidth     =   5100
   OleObjectBlob   =   "frmInputs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
'Close box and return to spreadsheet

Unload Me
Worksheets("Main").Activate

End Sub

Private Sub cmdClearForm_Click()
'Clears entries for all the text boxes on the form

txtTermServiceIndex = ""
txtRegionalFactor = ""
txtSoilSupport = ""

txtTargetESAL = ""


End Sub

Private Sub cmdCalculate_Click()
'Takes data from form, uses time interval and places hourly traffic flows on spreadsheet

'Declare variable types and names
Dim Pt As Single
Dim R As Integer
Dim S As Integer
Dim SN_new As Double
Dim SN_old As Double
Dim epsilon As Single
Dim ESAL(52) As Double
Dim SN(52) As Double

'Get user inputs for Pt, R, and S
Pt = .txt.TermServiceIndex.Value
R = .txtRegionalFactor.Value
S = .txt.SoilSupport.Value


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

'Initialize variables
maxESAL = 30000000
minESAL = 300000
lmin = Log(minESAL)
lmax = Log(maxESAL)

'Calculate structural numbers with corresponding ESALs
For I = 1 To 51
    ESAL(I) = Exp(lmin + (lmax - lmin) * 0.02 * (I - 1))
        
        epsilon = 0.001
        SN_new = 4
        SN_old = 0
        
        Do While Abs(SN_old - SN_new) > epsilon
            SN_old = SN_new
            SN_new = 10 ^ ((Log(ESAL(I)) / Log(10) + 0.2 - (Log((4.2 - Pt) / (4.2 - 1.5)) / Log(10)) / (0.4 + (1094 / (SN_old + 1) ^ 5.19)) + Log(R) / Log(10) - 0.372 * (S - 3)) / 9.36) - 1
        Loop
    
    SN(I) = SN_new
    Sheet2.Cells(I + 1, 1).Value = ESAL(I)
    Sheet2.Cells(I + 1, 2).Value = SN(I)
Next

End Sub
