' Fractal1

Option Explicit
Sub Main ()

	BeginHide
	StoreDoubleParameter "fractal_depth",   4
	StoreDoubleParameter "fractal_length", 10
	StoreDoubleParameter "fractal_width", 0.1
	EndHide

	Const MaxPoints = 10000

	Dim fractalDepth  As Integer
	Dim fractalLength As Double
	Dim fractalWidth  As Double
	Dim pointsX(MaxPoints) As Double
	Dim pointsY(MaxPoints) As Double
	Dim subdivisionX(MaxPoints) As Double
	Dim subdivisionY(MaxPoints) As Double
	Dim subdivisionCount As Integer
	Dim subdivisionPointIndex As Integer
	Dim currentCount As Integer
	Dim depthLevel As Integer

	Dim startPointX As Double
	Dim startPointY As Double
	Dim endPointX As Double
	Dim endPointY As Double
	Dim endPointIndex As Integer
	Dim distance As Double
	Dim cosF As Double
	Dim sinF As Double

	Dim i As Integer
	Dim j As Integer
	Dim x As Double
	Dim y As Double
	Dim xx As Double
	Dim yy As Double

	'
	fractalDepth  = CInt(RestoreDoubleParameter("fractal_depth"))
	fractalLength = RestoreDoubleParameter("fractal_length")
	fractalWidth  = RestoreDoubleParameter("fractal_width")

	'
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

			'
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

				x = x * distance
				y = y * distance

				xx = x*cosF + y*sinF
				yy = x*sinF - y*cosF

				xx = xx + startPointX
				yy = yy + startPointY

				subdivisionX(subdivisionPointIndex) = xx
				subdivisionY(subdivisionPointIndex) = yy
			Next

			'
			For i=MaxPoints-subdivisionCount To endPointIndex STEP -1
				pointsX(i+subdivisionCount) = pointsX(i)
				pointsY(i+subdivisionCount) = pointsY(i)
			Next

			For i=endPointIndex To endPointIndex+subdivisionCount-1
				pointsX(i) = subdivisionX(i-endPointIndex+1)
				pointsY(i) = subdivisionY(i-endPointIndex+1)
			Next

			currentCount = currentCount + subdivisionCount
		Next
	Next

	Curve.NewCurve "3D-Analytical"
	With Polygon3D
		.Reset
		.Name "fractal_curve"
		.Curve "3D-Analytical"

		For i = 1 To currentCount
			.Point pointsX(i), pointsY(i), 0
		Next
		.Create
	End With
End Sub

