VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBooster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private c As XlCalculation
Private s As Boolean
Private e As Boolean

Private Sub Class_Initialize()
    ' c = Application.Calculation
    s = Application.ScreenUpdating
    e = Application.EnableEvents

    'Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Application.Calculation = c
    Application.ScreenUpdating = s
    Application.EnableEvents = e
End Sub
