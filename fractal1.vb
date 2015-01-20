' Fractal1

Public Function CalcSegmentsNumber(ByVal N As Integer) As Integer
	CalcSegmentsNumber = 4^(N-1)
End Function

Public Function SubdivideSegment(ByRef Xs() As Double, ByRef Ys() As Double, ByVal i As Integer) As Integer

	Xs(i) = 1/3
	Ys(i) = 0
	i = i + 1

	Xs(i) = 1/2
	Ys(i) = Sqr(3)/6
	i = i + 1

	Xs(i) = 2/3
	Ys(i) = 0
	i = i + 1

	SubdivideSegment = 3
End Function

Public Sub InsertArray(ByRef target() As Double, ByVal index As Integer, ByVal source() As Double, ByVal count As Integer)
	Dim i As Integer
	Dim N As Integer
	N = UBound(target)
	For i=N-count To index STEP -1
		target(i+count) = target(i)
	Next

	For i=index To index+count-1
		target(i) = source(i-index+1)
	Next
End Sub


Option Explicit
Sub Main ()

	BeginHide
	StoreDoubleParameter "fractal_depth",   4
	StoreDoubleParameter "fractal_length", 10
	StoreDoubleParameter "fractal_width", 0.1
	EndHide

	Dim N As Integer
	Dim L As Double
	Dim D As Double
	N = CInt(RestoreDoubleParameter("fractal_depth"))
	L = RestoreDoubleParameter("fractal_length")
	D = RestoreDoubleParameter("fractal_width")

	Dim Q As Integer
	Dim P As Integer
	Q = CalcSegmentsNumber(N)
	P = Q + 1

	Dim Xs() As Double
	Dim Ys() As Double
	ReDim Xs(P)
	ReDim Ys(P)
	Xs(1) = 0
	Ys(1) = 0
	Xs(2) = L
	Ys(2) = 0

	Dim TempXs() As Double
	Dim TempYs() As Double
	ReDim TempXs(P)
	ReDim TempYs(P)

	Dim TempCount As Integer
	Dim level As Integer
	Dim i As Integer
	Dim j As Integer
	Dim c As Integer
	Dim x As Double
	Dim y As Double
	Dim xx As Double
	Dim yy As Double
	Dim x0 As Double
	Dim y0 As Double
	Dim x1 As Double
	Dim y1 As Double
	Dim dist As Double
	Dim cosF As Double
	Dim sinF As Double


	c = 2
	For level = 0 To N

		For i=c To 2 STEP -1

			x0 = Xs(i-1)
			y0 = Ys(i-1)
			x1 = Xs(i)
			y1 = Ys(i)
			dist = Sqr((x1-x0)^2 + (y1-y0)^2)
			cosF = (x1-x0) / dist
			sinF = (y1-y0) / dist

			TempCount = SubdivideSegment(TempXs, TempYs, 1)

			For j=1 To TempCount
				x = TempXs(j)
				y = TempYs(j)

				x = x * dist
				y = y * dist

				xx = x*cosF + y*sinF
				yy = x*sinF - y*cosF

				xx = xx + x0
				yy = yy + y0

				TempXs(j) = xx
				TempYs(j) = yy
			Next

			InsertArray(Xs, i, TempXs, TempCount)
			InsertArray(Ys, i, TempYs, TempCount)
		Next

		c = c + TempCount
	Next

	Curve.NewCurve "3D-Analytical"
	With Polygon3D
		.Reset
		.Name "fractal_curve"
		.Curve "3D-Analytical"

		For i = 0 To P
			.Point Xs(i), Ys(i), 0
		Next
		.Create
	End With

End Sub

