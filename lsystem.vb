' L-system

Option Explicit
Sub Main () 
	'
	BeginHide
	StoreDoubleParameter "lsystem_depth",   4
	StoreDoubleParameter "lsystem_width",   4
	StoreDoubleParameter "lsystem_height",  4
	StoreDoubleParameter "lsystem_step",    1
	StoreDoubleParameter "lsystem_angle_n", 6
	StoreParameter       "lsystem_axiom",   ""
	StoreParameter       "lsystem_rules",   ""
	SetParameterDescription("lsystem_axiom", "F")
	SetParameterDescription("lsystem_rules", "F = F-F++F-F")
	EndHide

	'
	Dim lsystemDepth As Integer
	Dim lsystemWidth As Double
	Dim lsystemHeight As Double
	Dim lsystemStep As Double
	Dim lsystemAngleN As Integer
	Dim lsystemAxiom As String
	Dim lsystemRules As String
	Dim ruleItem  As String
	Dim ruleComponents() As String
	Dim programSource As String
	Dim programTarget As String
	Dim iteration As Integer
	Dim symbol As String
	Dim i As Integer
	Dim x As Double
	Dim y As Double
	Dim angleN As Integer
	Dim angle As Double
	Dim curveIndex As Integer

	'
	Dim rulesHash
	Set rulesHash = CreateObject("Scripting.Dictionary")

	'
	Dim stackTip As Integer
	Dim stackHashX
	Dim stackHashY
	Dim stackHashA
	Set stackHashX = CreateObject("Scripting.Dictionary")
	Set stackHashY = CreateObject("Scripting.Dictionary")
	Set stackHashA = CreateObject("Scripting.Dictionary")
	lsystemDepth  = CInt(RestoreDoubleParameter("lsystem_depth"))
	lsystemWidth  = RestoreDoubleParameter("lsystem_width")
	lsystemHeight = RestoreDoubleParameter("lsystem_height")
	lsystemStep   = RestoreDoubleParameter("lsystem_step")
	lsystemAngleN = CInt(RestoreDoubleParameter("lsystem_angle_n"))
	lsystemAxiom  = GetParameterDescription("lsystem_axiom")
	lsystemRules  = GetParameterDescription("lsystem_rules")

	'
	lsystemAxiom = Replace(lsystemAxiom, " ", "")
	lsystemRules = Replace(lsystemRules, " ", "")

	For Each ruleItem In Split(lsystemRules, ";")
		ruleComponents = Split(ruleItem, "=")
		rulesHash.Add ruleComponents(0), ruleComponents(1)
	Next

	'
	programSource = lsystemAxiom
	For iteration = 1 To lsystemDepth
		For i = 1 To Len(programSource)
			symbol = Mid(programSource, i, 1)

			If rulesHash.Exists(symbol) Then
				programTarget = programTarget & rulesHash(symbol)
			Else
				programTarget = programTarget & symbol
			End If
		Next

		programSource = programTarget
		programTarget = ""
	Next

	x = 0
	y = 0
	angleN = 0
	angle  = 0
	curveIndex = 0
	stackTip = 0
	For i = 1 To Len(programSource)
		symbol = Mid(programSource, i, 1)

		If symbol = "F" Then
			curveIndex = curveIndex + 1
			Curve.NewCurve "3D-Analytical"
			Polygon3D.Reset
			Polygon3D.Name  "lsystem_curve" & curveIndex
			Polygon3D.Curve "3D-Analytical"
			Polygon3D.Point x, y, 0

			x = x + lsystemStep * Cos(angle)
			y = y + lsystemStep * Sin(angle)

			Polygon3D.Point x, y, 0
			Polygon3D.Create

			With TraceFromCurve
				.Reset
				.Name "lsystem_solid" & curveIndex
				.Component "component1"
				.Material "Vacuum"
				.Curve "3D-Analytical:lsystem_curve" & curveIndex
				.Thickness "0.1"
				.Width "0.1"
				.RoundStart "True"
				.RoundEnd "True"
				.GapType "2"
				.Create
			End With
			


		ElseIf symbol = "f" Then
			x = x + lsystemStep * Cos(angle)
			y = y + lsystemStep * Sin(angle)
		ElseIf symbol = "+" Then
			angleN = angleN + 1
			angle = 2*PI * angleN / lsystemAngleN
		ElseIf symbol = "-" Then
			angleN = angleN - 1
			angle = 2*PI * angleN / lsystemAngleN
		ElseIf symbol = "[" Then
			stackHashX(stackTip) = x
			stackHashY(stackTip) = y
			stackHashA(stackTip) = angleN
			stackTip = stackTip + 1

		ElseIf symbol = "]" Then
			stackTip = stackTip - 1
			x = stackHashX(stackTip)
			y = stackHashY(stackTip)
			angleN = stackHashA(stackTip)
		End If
	Next

	If curveIndex > 1 Then
		For i = 2 To curveIndex
			Solid.Add "component1:lsystem_solid1", "component1:lsystem_solid" & i
		Next
	End If
End Sub
