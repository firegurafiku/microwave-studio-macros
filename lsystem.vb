' L-system

Option Explicit
Sub Main()
	BeginHide
	StoreDoubleParameter "lsystem_depth",   4
	StoreDoubleParameter "lsystem_width",   4
	StoreDoubleParameter "lsystem_height",  4
	StoreDoubleParameter "lsystem_step",    1
	StoreDoubleParameter "lsystem_angle_n", 6
	StoreDoubleParameter "lsystem_axiom",   0
	StoreDoubleParameter "lsystem_rules",   0

	' Thank you CST for leaving me nothing but...
	SetParameterDescription "lsystem_axiom", "F"
	SetParameterDescription "lsystem_rules", "F = F-F++F-F"
	EndHide

	Dim lsystemDepth     As Integer
	Dim lsystemWidth     As Double
	Dim lsystemHeight    As Double
	Dim lsystemStep      As Double
	Dim lsystemAngleN    As Integer
	Dim lsystemAxiom     As String
	Dim lsystemRules     As String
	Dim ruleItem         As String
	Dim ruleComponents() As String
	Dim programSource    As String
	Dim programTarget    As String
	Dim iteration        As Integer
	Dim symbol           As String
	Dim curveIndex       As Integer
	Dim tortoiseX        As Double
	Dim tortoiseY        As Double
	Dim tortoiseA        As Integer
	Dim tortoiseAngle    As Double

	' Temporary variables.
	Dim i                As Integer

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

	'
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

	curveIndex = 0
	stackTip   = 0
	tortoiseX  = 0
	tortoiseY  = 0
	tortoiseA  = 0
	tortoiseAngle = 0
	For i = 1 To Len(programSource)
		symbol = Mid(programSource, i, 1)

		If symbol = "F" Then
			curveIndex = curveIndex + 1
			Curve.NewCurve "3D-Analytical"
			Polygon3D.Reset
			Polygon3D.Name  "curve" & curveIndex
			Polygon3D.Curve "3D-Analytical"
			Polygon3D.Point tortoiseX, tortoiseY, 0

			tortoiseAngle = 2*PI * tortoiseA / lsystemAngleN
			tortoiseX = tortoiseX + lsystemStep * Cos(tortoiseAngle)
			tortoiseY = tortoiseY + lsystemStep * Sin(tortoiseAngle)

			Polygon3D.Point tortoiseX, tortoiseY, 0
			Polygon3D.Create

			With TraceFromCurve
				.Reset
				.Name       "solid" & curveIndex
				.Component  "lsystem"
				.Material   "Vacuum"
				.Curve      "3D-Analytical:curve" & curveIndex
				.Thickness  lsystemHeight
				.Width      lsystemWidth
				.RoundStart True
				.RoundEnd   True
				.GapType    2
				.Create
			End With
		ElseIf symbol = "f" Then
			tortoiseAngle = 2*PI * tortoiseA / lsystemAngleN
			tortoiseX = tortoiseX + lsystemStep * Cos(tortoiseAngle)
			tortoiseY = tortoiseY + lsystemStep * Sin(tortoiseAngle)
		ElseIf symbol = "+" Then
			tortoiseA = tortoiseA + 1
		ElseIf symbol = "-" Then
			tortoiseA = tortoiseA - 1
		ElseIf symbol = "[" Then
			stackHashX(stackTip) = tortoiseX
			stackHashY(stackTip) = tortoiseY
			stackHashA(stackTip) = tortoiseA
			stackTip = stackTip + 1
		ElseIf symbol = "]" Then
			stackTip = stackTip - 1
			tortoiseX = stackHashX(stackTip)
			tortoiseY = stackHashY(stackTip)
			tortoiseA = stackHashA(stackTip)
		End If
	Next

	If curveIndex > 1 Then
		For i = 2 To curveIndex
			Solid.Add "lsystem:solid1", "lsystem:solid" & i
		Next
	End If
End Sub

