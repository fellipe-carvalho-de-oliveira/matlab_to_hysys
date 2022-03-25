VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Saturation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim hyContainer As HYSYS.ExtnUnitOperationContainer
Dim hyFeedStrm As HYSYS.ProcessStream
Dim hyProdStrm As HYSYS.ProcessStream
Dim Eq_of_State As Object
Dim Association_Model As Object
Dim W_fraction As Object
Dim CROSSH2S As Object
Dim CROSSCO2 As Object

Dim Components As Object
Dim Component As Object
Dim WorkFluid As Object
Dim Streams(0 To 1) As Object
Dim myOp As Object

Dim WaterPresent As Boolean
Dim H2O As Integer
Dim H2S As Integer
Dim N2 As Integer
Dim CO2 As Integer
Dim C1 As Integer
Dim C2 As Integer
Dim C3 As Integer
Dim C4 As Integer
Dim iC4 As Integer
Dim C5 As Integer
Dim iC5 As Integer
Dim C6 As Integer
Dim C7 As Integer
Dim C8 As Integer
Dim C9 As Integer
Dim C10 As Integer
Dim Temp As Variant
Dim MoleFl As Variant
Dim Count As Integer
Dim i As Integer
Dim IsKnownFeedArray As Variant
Dim CalcError(0 To 1) As Boolean
Dim IsIgnoredBool As Boolean
Const conEmpty As Integer = -32767

Public Function Initialize(ByVal Container As Object, ByVal IsRecalling As Boolean) As Long

On Error GoTo ErrorTrap

CalcError(0) = False
CalcError(1) = False
Initialize = extnCurrentVersion
Set hyContainer = Container
Set myOp = hyContainer.ExtensionInterface
Set Components = hyContainer.Flowsheet.FluidPackage.Components
WaterPresent = False
For Each Component In Components
  If Component.Name = "H2O" Then
    WaterPresent = True
  End If
Next
If WaterPresent Then H2O = Components.Index("H2O")
Set hyFeedStrm = hyContainer.FindVariable("FeedStream").Variable.object
Set hyProdStrm = hyContainer.FindVariable("ProductStream").Variable.object
Set Eq_of_State = hyContainer.FindVariable("Eos_Model").Variable
Set Association_Model = hyContainer.FindVariable("Assoc_Model").Variable
Set W_fraction = hyContainer.FindVariable("Water_Fraction").Variable
Set CROSSH2S = hyContainer.FindVariable("Cross_H2S_sol").Variable
Set CROSSCO2 = hyContainer.FindVariable("Cross_CO2_sol").Variable
If IsRecalling = False Then
  W_fraction.Value = 0
End If
ErrorTrap:

End Function

Public Sub Execute(ByVal Forgetting As Boolean)

On Error GoTo ErrorTrap
If Not Forgetting Then
  Set Components = hyContainer.Flowsheet.FluidPackage.Components
  WaterPresent = False
  For Each Component In Components
    If Component.Name = "H2O" Then
      WaterPresent = True
    End If
  Next Component
  If Not WaterPresent Then GoTo ErrorTrap
  H2O = Components.Index("H2O")
  H2S = Components.Index("H2S")
  N2 = Components.Index("Nitrogen")
  CO2 = Components.Index("CO2")
  C1 = Components.Index("Methane")
  C2 = Components.Index("Ethane")
  C3 = Components.Index("Propane")
  C4 = Components.Index("n-Butane")
  iC4 = Components.Index("i-Butane")
  C5 = Components.Index("n-Pentane")
  iC5 = Components.Index("i-Pentane")
  C6 = Components.Index("n-Hexane")
  C7 = Components.Index("n-Heptane")
  C8 = Components.Index("n-Octane")
  C9 = Components.Index("n-Nonane")
  C10 = Components.Index("n-Decane")
  Set hyFeedStrm = hyContainer.FindVariable("FeedStream").Variable.object
  Set hyProdStrm = hyContainer.FindVariable("ProductStream").Variable.object
  Set Eq_of_State = hyContainer.FindVariable("Eos_Model").Variable
  Set Association_Model = hyContainer.FindVariable("Assoc_Model").Variable
  Set W_fraction = hyContainer.FindVariable("Water_Fraction").Variable
  Set CROSSH2S = hyContainer.FindVariable("Cross_H2S_sol").Variable
  Set CROSSCO2 = hyContainer.FindVariable("Cross_CO2_sol").Variable
  If hyFeedStrm Is Nothing Then GoTo ErrorTrap
  If hyProdStrm Is Nothing Then GoTo ErrorTrap
  If hyFeedStrm.TemperatureValue = conEmpty Then GoTo ErrorTrap
  If hyFeedStrm.PressureValue = conEmpty Then GoTo ErrorTrap
  If hyFeedStrm.MolarFlowValue = conEmpty Then GoTo ErrorTrap
  Temp = hyFeedStrm.ComponentMolarFraction
  If (Temp(0) = conEmpty) Then GoTo ErrorTrap
  
  Set WorkFluid = hyFeedStrm.DuplicateFluid
  MoleFl = WorkFluid.MolarFractionsValue
  
  Dim Result As String
  Dim eos_model As Integer
  Dim assoc_model As Integer
  Dim Press As Double
  Dim Tempe As Double
  Dim cross_H2S As Integer
  Dim cross_CO2 As Integer
  Dim x_h2s As Double
  Dim x_n2 As Double
  Dim x_co2 As Double
  Dim x_c1 As Double
  Dim x_c2 As Double
  Dim x_c3 As Double
  Dim x_c4 As Double
  Dim x_ic4 As Double
  Dim x_c5 As Double
  Dim x_ic5 As Double
  Dim x_c6 As Double
  Dim x_c7 As Double
  Dim x_c8 As Double
  Dim x_c9 As Double
  Dim x_c10 As Double
  

  eos_model = Eq_of_State
  assoc_model = Association_Model
  Press = hyFeedStrm.PressureValue
  Tempe = hyFeedStrm.TemperatureValue
  cross_H2S = CROSSH2S
  cross_CO2 = CROSSCO2
  x_h2s = MoleFl(H2S)
  x_n2 = MoleFl(N2)
  x_co2 = MoleFl(CO2)
  x_c1 = MoleFl(C1)
  x_c2 = MoleFl(C2)
  x_c3 = MoleFl(C3)
  x_c4 = MoleFl(C4)
  x_ic4 = MoleFl(iC4)
  x_c5 = MoleFl(C5)
  x_ic5 = MoleFl(iC5)
  x_c6 = MoleFl(C6)
  x_c7 = MoleFl(C7)
  x_c8 = MoleFl(C8)
  x_c9 = MoleFl(C9)
  x_c10 = MoleFl(C10)
  
  Result = Shell(CurDir() + "\ATOMS_Saturation.exe" + MyFunction2(eos_model, assoc_model, Press, Tempe, cross_H2S, cross_CO2, x_h2s, x_n2, x_co2, x_c1, x_c2, x_c3, x_c4, x_ic4, x_c5, x_ic5, x_c6, x_c7, x_c8, x_c9, x_c10))
  
  'Pausing
  Dim PauseTime, Start, Finish, TotalTime
  PauseTime = 10   ' Set duration.
  Start = Timer   ' Set start time.
  Do While Timer < Start + PauseTime
    DoEvents    ' Yield to other processes.
  Loop
  Finish = Timer  ' Set end time.
  TotalTime = Finish - Start  ' Calculate total time.

  
  Dim water_frac As String
  Open "file.txt" For Input As #1   ' Open file for input.
    Line Input #1, water_frac    ' Read line of data.
  Close #1
  

  
  Dim Files As String
  Dim fs As Object
  
  Files = CurDir() + "\file.txt"
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  fs.DeleteFile Files, True

  
EndCalcs:
 WorkFluid.Erase
  
  With hyProdStrm
    .Pressure.Erase
    .Pressure.Calculate hyFeedStrm.PressureValue
    .Temperature.Erase
    .Temperature.Calculate hyFeedStrm.TemperatureValue
  End With

  
  If hyFeedStrm.MolarFlowValue = 0 Then
    hyProdStrm.ComponentMolarFraction.Calculate hyFeedStrm.ComponentMolarFractionValue
    hyProdStrm.MolarFlow.Calculate hyFeedStrm.MolarFlowValue
  Else
    Set Streams(0) = hyFeedStrm
    Set Streams(1) = hyProdStrm
    hyContainer.Balance btMoleBalance, 1, Streams
  End If

  
  W_fraction.Value = CDbl(Val(water_frac)) * 1000000
  hyContainer.SolveComplete
  CalcError(0) = False
  CalcError(1) = False


End If
ErrorTrap:

End Sub

Public Function MyFunction2(ByVal eos_model As Integer, ByVal assoc_model As Integer, ByVal Press As Double, ByVal Tempe As Double, ByVal cross_H2S As Integer, ByVal cross_CO2 As Integer, ByVal x_h2s As Double, ByVal x_n2 As Double, ByVal x_co2 As Double, ByVal x_c1 As Double, ByVal x_c2 As Double, ByVal x_c3 As Double, ByVal x_c4 As Double, ByVal x_ic4 As Double, ByVal x_c5 As Double, ByVal x_ic5 As Double, ByVal x_c6 As Double, ByVal x_c7 As Double, ByVal x_c8 As Double, ByVal x_c9 As Double, ByVal x_c10 As Double) As String
  MyFunction2 = " " + Str(eos_model) + Str(assoc_model) + Str(Press) + Str(Tempe) + Str(cross_H2S) + Str(cross_CO2) + Str(x_h2s) + Str(x_n2) + Str(x_co2) + Str(x_c1) + Str(x_c2) + Str(x_c3) + Str(x_c4) + Str(x_ic4) + Str(x_c5) + Str(x_ic5) + Str(x_c6) + Str(x_c7) + Str(x_c8) + Str(x_c9) + Str(x_c10)
End Function

Public Sub StatusQuery(hyStatus As ObjectStatus)

Dim OK As Boolean
OK = True

If WaterPresent = False Then
  Call hyStatus.AddStatusCondition(slMissingRequiredInformation, 1, "Water is Required as a Component")
  OK = False
End If

If hyFeedStrm Is Nothing Then
  Call hyStatus.AddStatusCondition(slMissingRequiredInformation, 2, "Feed Stream Required")
  OK = False
End If

If hyProdStrm Is Nothing Then
  Call hyStatus.AddStatusCondition(slMissingRequiredInformation, 3, "Product Stream Required")
  OK = False
End If

If myOp.IsIgnored = True Then
  Call hyStatus.AddStatusCondition(slWarning, 11, "Ignored")
  OK = False
End If

If OK = False Then Exit Sub

If Not hyFeedStrm.Temperature.IsKnown Then
  Call hyStatus.AddStatusCondition(slMissingOptionalInformation, 5, "Unknown Feed Temperature")
  OK = False
End If

If Not hyFeedStrm.Pressure.IsKnown Then
  Call hyStatus.AddStatusCondition(slMissingOptionalInformation, 6, "Unknown Feed Pressure")
  OK = False
End If

If Not hyFeedStrm.MolarFlow.IsKnown Then
  Call hyStatus.AddStatusCondition(slMissingOptionalInformation, 7, "Unknown Feed Flow")
  OK = False
End If

IsKnownFeedArray = hyFeedStrm.ComponentMolarFraction.IsKnown
If Not IsKnownFeedArray(0) Then
  Call hyStatus.AddStatusCondition(slMissingOptionalInformation, 8, "Unknown Feed Composition")
  OK = False
End If

If CalcError(0) Then
  Call hyStatus.AddStatusCondition(slError, 9, "Feed Cannot be Saturated with Water")
  OK = False
End If

If CalcError(1) Then
  Call hyStatus.AddStatusCondition(slError, 10, "Water is required in Water Stream")
  OK = False
End If

End Sub


