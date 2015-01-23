' L-system
' *** Be careful with the line above as CST uses it as a macro name.
'
' This script constructs a planar fractal antenna in CST Microwave Studio.
' It uses L-system grammar to draw a fractal of arbitrary type. After being
' executed from the “Macros” menu, it creates component “lsystem” holding
' generated antenna's solid:
'
'   * lsystem_length
'   * lsystem_width
'   * lsystem_height
'   * lsystem_depth
'   * lsystem_angle_n
'   * lsystem_axiom
'   * lsystem_rules
'
' “Axiom”, “rules” and “angle_n” parameters are the ones which specify the
' exact type of fractal to paint. Their values are string, which are held in
' parameter DESCRIPTIONS, not the values itself. I'm very sorry for this more
' then bad design decision, but CST engineers left me no other option to deal
' with their software. Here rules are set of replacements consequently applied
' to initial axiom written like in the example below:
'
'     lsystem_angle_n = 4
'     lsystem_axiom   = "X"
'     lsystem_rules   = "X = -YF+XFX+FY-; Y = +XF-YFY-FX+"
'
' This example contains two rules describing Hilbert curve. “Angle_n“ is the
' angle step for “+” and “-” commands specified as 360°/n. Rules are separated
' by semicolon, symbol and replacement are separated by equality sign within
' each rule. All spaces are ignored.
'
' Copyright (c) Pavel Kretov, 2015.
' Provided under the terms of MIT license. Created as a part of PhD research
' in Voronezh State University, Russia.
Option Explicit
Sub Main()
	BeginHide
	StoreDoubleParameter "lsystem_depth",   4
	StoreDoubleParameter "lsystem_width",   0.1
	StoreDoubleParameter "lsystem_height",  0
	StoreDoubleParameter "lsystem_step",    1
	StoreDoubleParameter "lsystem_angle_n", 6
	StoreDoubleParameter "lsystem_axiom",   0
	StoreDoubleParameter "lsystem_rules",   0

	' Thank you CST for leaving me nothing but...
	SetParameterDescription "lsystem_axiom", "F"
	SetParameterDescription "lsystem_rules", "F = F-F++F-F"
	EndHide

	' Don't pretend Visual Basic for application can handle nested variable
	' scope and just declare *all* variables in a single place.
	Dim lsystemDepth     As Integer
	Dim lsystemWidth     As Double
	Dim lsystemHeight    As Double
	Dim lsystemStep      As Double
	Dim lsystemAngleN    As Integer
	Dim lsystemAxiom     As String
	Dim lsystemRules     As String
	Dim simpleRendering  As Boolean
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

	' Create a dictionary for rules. The “Scripting.Dictionary” here is
	' the name of COM object from Windows Script Host.
	Dim rulesHash
	Set rulesHash = CreateObject("Scripting.Dictionary")

	' Create dictionaries for stacks. Here stacks are used instead of
	' arrays to avoid dealing with memory issues or introducing .NET
	' dependency as in case of “System.Collections.ArrayList”.
	Dim stackTip As Integer
	Dim stackHashX
	Dim stackHashY
	Dim stackHashA
	Set stackHashX = CreateObject("Scripting.Dictionary")
	Set stackHashY = CreateObject("Scripting.Dictionary")
	Set stackHashA = CreateObject("Scripting.Dictionary")

	' Read global parameters.
	lsystemDepth  = CInt(RestoreDoubleParameter("lsystem_depth"))
	lsystemWidth  = RestoreDoubleParameter("lsystem_width")
	lsystemHeight = RestoreDoubleParameter("lsystem_height")
	lsystemStep   = RestoreDoubleParameter("lsystem_step")
	lsystemAngleN = CInt(RestoreDoubleParameter("lsystem_angle_n"))
	lsystemAxiom  = GetParameterDescription("lsystem_axiom")
	lsystemRules  = GetParameterDescription("lsystem_rules")

	' Strip all apaces.
	lsystemAxiom = Replace(lsystemAxiom, " ", "")
	lsystemRules = Replace(lsystemRules, " ", "")

	' Populate dictionaryu with rules for faster lookup.
	For Each ruleItem In Split(lsystemRules, ";")
		ruleComponents = Split(ruleItem, "=")
		rulesHash.Add ruleComponents(0), ruleComponents(1)
	Next

	' Generate program code for tortoise interpreter. Plain string
	' concatenation here is kind of slow, but very simple and doesn't
	' require including external dependencies.
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

	' Try to detect if fractal can be drawn with a single polyline. This
	' check is very simple but can output some false negatives.
	simpleRendering = True
	If InStr(1, programSource, "f") <> 0 Or _
	   InStr(1, programSource, "[") <> 0 Or _
	   InStr(1, programSource, "]") <> 0 Then
		simpleRendering = False
	End If

	' Use a tortoise interpreter with memory for actual fractal drawing.
	' Parameters below are initial state. The reason why instead of
	' floating angle we use integral 'angleN' is avoiding floating point
	' truncation errors (which are more then possible when drawing lines
	' with lots of turns left and right back, like fractals are).
	curveIndex = 0
	stackTip   = 0
	tortoiseX  = 0
	tortoiseY  = 0
	tortoiseA  = 0
	tortoiseAngle = 0
	If simpleRendering Then
		' If the fractal can be drawn in a single line, then just draw
		' the line completely, then convert it to solid. This is much
		' faster.
		Polygon3D.Reset
		Polygon3D.Curve "lsystem"
		Polygon3D.Name  "curve1"

		Polygon3D.Point tortoiseX, tortoiseY, 0
		For i = 1 To Len(programSource)
			symbol = Mid(programSource, i, 1)

			If symbol = "F" Then
				tortoiseAngle = 2*PI * tortoiseA / lsystemAngleN
				tortoiseX = tortoiseX + lsystemStep * Cos(tortoiseAngle)
				tortoiseY = tortoiseY + lsystemStep * Sin(tortoiseAngle)
				Polygon3D.Point tortoiseX, tortoiseY, 0
			ElseIf symbol = "+" Then
				tortoiseA = tortoiseA + 1
			ElseIf symbol = "-" Then
				tortoiseA = tortoiseA - 1
			End If
		Next

		Polygon3D.Create
		With TraceFromCurve
			.Reset
			.Name       "solid1"
			.Component  "lsystem"
			.Material   "Vacuum"
			.Curve      "lsystem:curve1"
			.Thickness  lsystemHeight
			.Width      lsystemWidth
			.RoundStart True
			.RoundEnd   True
			.GapType    2
			.Create
		End With
	Else
		' If the fractal is too complex to be drawn in a single line,
		' then there is nothing but to draw each line segment
		' separately, then convert it to solid, and finally combine
		' all the lines into a single solid object. Which is much,
		' much slower then the first approach.
		For i = 1 To Len(programSource)
			symbol = Mid(programSource, i, 1)

			If symbol = "F" Then
				curveIndex = curveIndex + 1
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
	End If

	If curveIndex > 1 Then
		For i = 2 To curveIndex
			Solid.Add "lsystem:solid1", "lsystem:solid" & i
		Next
	End If
End Sub
