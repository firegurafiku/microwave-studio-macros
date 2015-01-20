' Vivaldi
' !!! Do not change the line above !!!

Option Explicit
Sub Main ()

	BeginHide
	StoreDoubleParameter "vivaldi_length",   10.0
	StoreDoubleParameter "vivaldi_width",    10.0
	StoreDoubleParameter "vivaldi_height",    1.0
	StoreDoubleParameter "vivaldi_overlap",   1.0
	StoreDoubleParameter "vivaldi_aperture",  8.0
	StoreDoubleParameter "vivaldi_exponent",  1.0
	StoreDoubleParameter "vivaldi_nxpoints",   50
	EndHide

	Dim L, W, H, b, d, a, n, dx As Double
	L = RestoreDoubleParameter("vivaldi_length")
	W = RestoreDoubleParameter("vivaldi_width")
	H = RestoreDoubleParameter("vivaldi_height")
	b = RestoreDoubleParameter("vivaldi_overlap")
	d = RestoreDoubleParameter("vivaldi_aperture")
	a = RestoreDoubleParameter("vivaldi_exponent")
	n = RestoreDoubleParameter("vivaldi_nxpoints")
	dx = W / n

	Dim x, y As Double

	' Begin construction
	Curve.NewCurve "3D-Analytical"
	With Polygon3D
		.Reset
		.Name "vivaldi_curve1"
		.Curve "3D-Analytical"

		For x = 0 To L STEP dx
			y = b/2 - (1 - Exp(x^a)) * (b+d) / 2 / (1 - Exp(L^a))
			.Point x, y, H/2
		Next
		.Point L, -W/2, H/2
		.Point 0, -W/2, H/2
		.Point 0,  b/2, H/2
		.Create
	End With

	With CoverCurve
	.Reset
	.Name "vivaldi_solid1"
	.Component "vivaldi_component1"
	.Material "PEC"
	.Curve "3D-Analytical:vivaldi_curve1"
	.Create
	End With

	With Transform
	     .Reset
	     .Name "vivaldi_component1"
	     .Origin "Free"
	     .Center "0", "0", "0"
	     .Angle "180", "0", "0"
	     .MultipleObjects "True"
	     .GroupObjects "False"
	     .Repetitions "1"
	     .MultipleSelection "False"
	     .Destination ""
	     .Material ""
	     .Transform "Shape", "Rotate"
	End With

	With Brick
	     .Reset
	     .Name "vivaldi_substrate"
	     .Component "vivaldi_substrate_component1"
	     .Material "Vacuum"
	     .Xrange 0, L
	     .Yrange -W/2, W/2
	     .Zrange -H/2, H/2
	     .Create
	End With
End Sub

