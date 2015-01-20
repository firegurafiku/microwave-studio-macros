' Fractal1
' *** Be careful with the line above as CST uses it as a macro name.
'
' This script constructs a planar fractal antenna in CST Microwave Studio. It
' can paint an approximation of a polyline fractal with a constant segment
' subdivision procedure. The exact form of fractal is can be specified in the
' code only, sorry. This is due to CST limitations.
'
' When being executed from the 'Macros' menu, the script paints a curve with
' name 'fractal_curve' which form can be controlled by the following
' global parameters (with default values):
'
'  * fractal_depth  = 4
'  * fractal_length = 10
'
' The maximum numbers of points in antenna is constrained by 'MaxPoints'
' constant in the code. Adjust it if you want to get a very deep fractal
' approximation it CST. Also, get a *planar* antenna out of that curve, apply
' “Curves -> Trace from Curve...” transformation to it. It seems to work well
' for polylines.
'
' Copyright (c) Pavel Kretov, 2015.
' Provided under the terms of MIT license. This script was created as a part
' of PhD research in Voronezh State University, Russia.

Option Explicit
Sub Main ()
	BeginHide
	StoreDoubleParameter "fractal_depth",   4
	StoreDoubleParameter "fractal_length", 10
	EndHide

	' NOTE: Increase that number if you need to go deeper.
	Const MaxPoints = 10000

	' Don't pretend Visual Basic for application can handle nested variable
	' scope and just declare *all* variables in a single place.
	Dim fractalDepth            As Integer
	Dim fractalLength           As Double
	Dim fractalWidth            As Double
	Dim pointsX(MaxPoints)      As Double
	Dim pointsY(MaxPoints)      As Double
	Dim subdivisionX(MaxPoints) As Double
	Dim subdivisionY(MaxPoints) As Double
	Dim subdivisionCount        As Integer
	Dim subdivisionPointIndex   As Integer
	Dim currentCount            As Integer
	Dim depthLevel              As Integer
	Dim startPointX             As Double
	Dim startPointY             As Double
	Dim endPointX               As Double
	Dim endPointY               As Double
	Dim endPointIndex           As Integer
	Dim distance                As Double
	Dim cosF                    As Double
	Dim sinF                    As Double

	' Temporary variables. Don't assume any specific meaning for them.
	Dim i                       As Integer
	Dim j                       As Integer
	Dim x                       As Double
	Dim y                       As Double
	Dim xx                      As Double
	Dim yy                      As Double

	' Read global parameters.
	fractalDepth  = CInt(RestoreDoubleParameter("fractal_depth"))
	fractalLength =      RestoreDoubleParameter("fractal_length")

	' Put endpoints of the first segment.
	pointsX(1) = 0
	pointsY(1) = 0
	pointsX(2) = fractalLength
	pointsY(2) = 0

	currentCount = 2
	For depthLevel = 1 To fractalDepth
		For endPointIndex=currentCount To 2 STEP -1
			startPointX = pointsX(endPointIndex-1)
			startPointY = pointsY(endPointIndex-1)
			endPointX   = pointsX(endPointIndex)
			endPointY   = pointsY(endPointIndex)
			distance    = Sqr((endPointX-startPointX)^2 + (endPointY-startPointX)^2)
			cosF        = (endPointX-startPointX) / distance
			sinF        = (endPointY-startPointY) / distance

			' NOTE: Alter this block to construct another linear fractal.
			' To perform a subdivision, first, assign to 'subdivisionCount'
			' the number of points to *insert*. Then, put the coordinates
			' of these points to 'subdivisionX' and 'subdivisionY' arrays.
			subdivisionCount = 3
			subdivisionX(1) = 1/3
			subdivisionY(1) = 0
			subdivisionX(2) = 1/2
			subdivisionY(2) = Sqr(3)/6
			subdivisionX(3) = 2/3
			subdivisionY(3) = 0

			For subdivisionPointIndex=1 To subdivisionCount
				x = subdivisionX(subdivisionPointIndex)
				y = subdivisionY(subdivisionPointIndex)

				' First, scale ...
				x = x * distance
				y = y * distance

				' Then, rotate ...
				xx = x*cosF + y*sinF
				yy = x*sinF - y*cosF

				' Than, shift. In that order.
				xx = xx + startPointX
				yy = yy + startPointY

				subdivisionX(subdivisionPointIndex) = xx
				subdivisionY(subdivisionPointIndex) = yy
			Next

			' Make enough place by shifting part of items to the end.
			For i=MaxPoints-subdivisionCount To endPointIndex STEP -1
				pointsX(i+subdivisionCount) = pointsX(i)
				pointsY(i+subdivisionCount) = pointsY(i)
			Next

			' Insert new points into list.
			For i=endPointIndex To endPointIndex+subdivisionCount-1
				pointsX(i) = subdivisionX(i-endPointIndex+1)
				pointsY(i) = subdivisionY(i-endPointIndex+1)
			Next

			' Count new points just added.
			currentCount = currentCount + subdivisionCount
		Next
	Next

	' Finally, just draw a polyline curve using points which were calculated
	' at the previous step.
	Curve.NewCurve "3D-Analytical"
	With Polygon3D
		.Reset
		.Name  "fractal_curve"
		.Curve "3D-Analytical"
		For i = 1 To currentCount
			.Point pointsX(i), pointsY(i), 0
		Next
		.Create
	End With
End Sub

