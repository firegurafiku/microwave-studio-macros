' VSU\VarEpsilon

Option Explicit
Sub Main()
    BeginHide
    StoreDoubleParameter "varepsilon_Nepsilon", 100
    StoreDoubleParameter "varepsilon_dx",       0.1
    StoreDoubleParameter "varepsilon_dy",       0.1
    StoreDoubleParameter "varepsilon_dz",       0.1
    StoreDoubleParameter "varepsilon_Lx",       10
    StoreDoubleParameter "varepsilon_Ly",       10
    StoreDoubleParameter "varepsilon_Lz",       10
    StoreDoubleParameter "varepsilon_x0",       0
    StoreDoubleParameter "varepsilon_y0",       0
    StoreDoubleParameter "varepsilon_z0",       0
    EndHide

    Dim Nepsilon As Double
    Nepsilon = RestoreDoubleParameter("varepsilon_Nepsilon")

    Dim dx, dy, dz As Double
    Dim Lx, Ly, Lz As Double
    Dim x0, y0, z0 As Double
    dx = RestoreDoubleParameter("varepsilon_dx")
    dy = RestoreDoubleParameter("varepsilon_dy")
    dz = RestoreDoubleParameter("varepsilon_dz")
    Lx = RestoreDoubleParameter("varepsilon_Lx")
    Ly = RestoreDoubleParameter("varepsilon_Ly")
    Lz = RestoreDoubleParameter("varepsilon_Lz")
    x0 = RestoreDoubleParameter("varepsilon_x0")
    y0 = RestoreDoubleParameter("varepsilon_y0")
    z0 = RestoreDoubleParameter("varepsilon_z0")

    Dim Nx, Ny, Nz As Integer
    Nx = Int(Lx / dx)
    Ny = Int(Ly / dy)
    Nz = Int(Lz / dz)

    ReDim epsilonSums(Nx, Ny, Nz)  As Double
    ReDim pointCounts(Nx, Ny, Nz)  As Double
    Dim minimumEpsilon             As Double
    Dim maximumEpsilon             As Double

    ' HACK: That's not the right way to deal with things in general, but our
    '       specific case.
    minimumEpsilon = 1E6
    maximumEpsilon = 0

    Dim sLine    As String
    Dim sArr3(4) As Double
    Dim sArr

    Dim i, j As Integer
    Dim ix, iy, iz As Integer
    Dim xc, yc, zc, ec As Double

    Open "varepsilon_data.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, sLine
        sArr = Split(sLine)

        j = 1
        For i = LBound(sArr) To UBound(sArr)
            Dim s As String
            s = sArr(i)
            If s <> "" Then
                sArr3(j) = Val(s)
                j = j + 1
            End If
        Next

        If j <> 5 Then
            MsgBox "HUITA"
        End If

        xc = sArr3(1) + x0
        yc = sArr3(2) + y0
        zc = sArr3(3) + z0
        ec = sArr3(4)

        ix = Int(xc / dx)
        iy = Int(yc / dy)
        iz = Int(zc / dz)
        pointCounts(ix, iy, iz) = pointCounts(ix, iy, iz) + 1
        epsilonSums(ix, iy, iz) = epsilonSums(ix, iy, iz) + ec

        If ec < minimumEpsilon Then
            minimumEpsilon = ec
        End If

        If ec > maximumEpsilon Then
            maximumEpsilon = ec
        End If
    Wend
    Close #1

    Dim cnt As Integer
    Dim sum As Double

    Dim eps As Double
    Dim epsIdx As Integer

    Dim depsilon As Double
    depsilon = (maximumEpsilon - minimumEpsilon) / Nepsilon

    ReDim epsilons(Nepsilon) As Double

    For ix = 0 To Nx-1
        For iy = 0 To Ny-1
            For iz = 0 To Nz-1
                cnt = pointCounts(ix, iy, iz)
                sum = epsilonSums(ix, iy, iz)

                If cnt <> 0 Then
                    eps = sum / cnt
                    epsIdx = Int((eps - minimumEpsilon) / (maximumEpsilon - minimumEpsilon) * Nepsilon)
                    epsilons(epsIdx) = epsilons(epsIdx) + 1
                End If
            Next
        Next
    Next

    For i = LBound(epsilons) To UBound(epsilons)
        If epsilons(i) <> 0 Then
        With Material
             .Reset
             .Name                      "material_" & i
             .Folder                    "varepsilon"
             .Epsilon                   minimumEpsilon + i*depsilon
             .Transparency              0
             .FrqType                   "all"
             .Type                      "Normal"
             .SetMaterialUnit           "Hz", "m"
             ' ---
             .Mue                       1
             .ThinPanel                 False
             .Thickness                 0
             .Kappa                     0
             .TanD                      0.0
             .TanDFreq                  0.0
             .TanDGiven                 False
             .TanDModel                 "ConstTanD"
             .ConstTanDModelOrderEps     1
             .ReferenceCoordSystem       "Global"
             .CoordSystemType            "Cartesian"
             .KappaM                     0
             .TanDM                      0.0
             .TanDMFreq                  0.0
             .TanDMGiven                 False
             .TanDMModel                 "ConstTanD"
             .ConstTanDModelOrderMue     1
             .DispModelEps               "None"
             .DispModelMue               "None"
             .DispersiveFittingSchemeEps "1st Order"
             .DispersiveFittingSchemeMue "1st Order"
             .UseGeneralDispersionEps    False
             .UseGeneralDispersionMue    False
             .NLAnisotropy               False
             .NLAStackingFactor          1
             .NLADirectionX              1
             .NLADirectionY              0
             .NLADirectionZ              0
             .Rho                        0
             .ThermalType                "Normal"
             .ThermalConductivity        0
             .HeatCapacity               0
             .MetabolicRate              0
             .BloodFlow                  0
             .VoxelConvection            0
             .MechanicsType              "Unused"
             .Colour                     0, 1, 1
             .Wireframe                  False
             .Reflection                 False
             .Allowoutline               True
             .Transparentoutline         False
             .Create
        End With
        End If
    Next

    For ix = 0 To Nx-1
        For iy = 0 To Ny-1
            For iz = 0 To Nz-1
                cnt = pointCounts(ix, iy, iz)
                sum = epsilonSums(ix, iy, iz)

                If cnt <> 0 Then
                    eps = sum / cnt
                    epsIdx = Int((eps - minimumEpsilon) / (maximumEpsilon - minimumEpsilon) * Nepsilon)
                    With Brick
                         .Reset
                         .Name "solid_" & ix & "_" & iy & "_" & iz
                         .Component "varepsilon"
                         .Material "varepsilon/material_" & epsIdx
                         .Xrange ix*dx, (ix+1)*dx
                         .Yrange iy*dy, (iy+1)*dy
                         .Zrange iz*dz, (iz+1)*dz
                         .Create
                    End With

                With Transform
                     .Reset
                     .Name "varepsilon:cutoff"
                     .Vector "0", "0", "0"
                     .UsePickedPoints "False"
                     .InvertPickedPoints "False"
                     .MultipleObjects "True"
                     .GroupObjects "False"
                     .Repetitions "1"
                     .MultipleSelection "False"
                     .Destination ""
                     .Material ""
                     .Transform "Shape", "Translate"
                End With

                With Solid
                 .Version 9
                 .Intersect "varepsilon:solid_" & ix & "_" & iy & "_" & iz, "varepsilon:cutoff_1"
                 .Version 1
                End With

                End If
            Next
        Next
    Next
End Sub

