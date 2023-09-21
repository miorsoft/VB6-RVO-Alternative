Attribute VB_Name = "mRVO"
Option Explicit


Public Type tAgent
    X             As Double
    Y             As Double
    VX            As Double
    VY            As Double
    TX            As Double
    TY            As Double
    VXchange      As Double
    VYchange      As Double

    cR            As Double
    cG            As Double
    cB            As Double

    V             As Double

    SciaX()       As Double
    SciaY()       As Double

    ShoulderX     As Double
    ShoulderY     As Double

    Avoidance         As Double

    NReachedTargets As Long

End Type

Private TRAILS      As Long

Public Const Ntrail As Long = 1    ' 6    '4

Public Agnet()    As tAgent
Public NA         As Long

Public WorldW     As Double
Public WorldH     As Double

Private SRF       As cCairoSurface
Private CC        As cCairoContext

Private Const MaxV As Double = 1  '0.52
Private Const MaxV2 As Double = MaxV * MaxV
Private Const InvMaxV As Double = 1 / MaxV

Private GRID      As cSpatialGrid

Public doloop     As Boolean

Private CNT       As Long

Private Const ControlDist As Double = 14    ' 14 ' 14 ' 14    '12
Private Const ControlDist2 As Double = ControlDist * ControlDist
Private Const invControlDist As Double = 1 / ControlDist

Public Const VMUL As Double = 17  ' 20 '20  '30 ' 25 '28

Private Const GridSize As Double = 2 * MaxV * VMUL + ControlDist    '40    '35 '25

Private Const VelChangeGlobStrngth As Double = 0.2    'con MaxV

Public MODE       As Long
Public FOLLOW     As Long

Private Const ZOOM As Double = 1.3    '1.4
Private Const InvZOOM As Double = 1 / ZOOM

Private Const PI  As Single = 3.14159265358979
Private Const PIh As Single = 1.5707963267949
Private Const PI2 As Single = 6.28318530717959
Private Const InvPI As Single = 1 / 3.14159265358979
Private Const InvPIh As Single = 1 / 1.5707963267949
Private Const InvPI2 As Single = 1 / 6.28318530717959

Private Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
    If X Then
        '        Atan2 = -PI + Atn(Y / X) - (X > 0#) * Pi
        Atan2 = Atn(Y / X) + PI * (X < 0#)
    Else
        Atan2 = -PIh - (Y > 0#) * PI
    End If
    ' While Atan2 < 0: Atan2 = Atan2 + PI2: Wend
    ' While Atan2 > PI2: Atan2 = Atan2 - PI2: Wend
End Function


Public Sub Init_RVO(NAG As Long)
    Dim I         As Long
    Dim ANG#

    Const minAvoid  As Double = 0.2
    Const OneMinusMinAvoid As Double = 1 - minAvoid


    NA = NAG
    ReDim Agnet(NA)
    For I = 1 To NA

        With Agnet(I)
            .X = (0.1 + 0.8 * Rnd) * WorldW
            .Y = (0.1 + 0.8 * Rnd) * WorldH
            .TX = (0.1 + 0.8 * Rnd) * WorldW
            .TY = (0.1 + 0.8 * Rnd) * WorldH

            .Avoidance = minAvoid + OneMinusMinAvoid * Rnd

            ReDim .SciaX(Ntrail)
            ReDim .SciaY(Ntrail)


            .cG = (.Avoidance - minAvoid) / (OneMinusMinAvoid): .cB = 0.3
            .cR = 1 - .cG

            '            Do
            '                .cR = 1 - .Avoidance: .cG = .Avoidance: .cB = Rnd
            '
            '            Loop While .cR * 3 + .cG * 5 + .cB * 2 < 4

            ''''             CIRLCE
            ''            Ang = I / NA * 6.28
            ''            .TX = WorldW * 0.5 + Cos(Ang) * WorldW * 0.4
            ''            .TY = WorldH * 0.5 + Sin(Ang) * WorldH * 0.4
            ''
            ''.X = .TX
            ''.Y = .TY
            ''
            ''.cR = 0.6 + 0.4 * Cos(Ang * 2 + 3)
            ''.cG = 0.6 + 0.4 * Cos(Ang * 1 + 0.7)
            ''.cB = 0.6 + 0.4 * Cos(Ang * 2 + 0.5)
            ''
            '''            .cR = .cR - Int(.cR) + 0.1
            '''            .cG = .cG - Int(.cG) + 0.1
            '''            .cB = .cB - Int(.cB) + 0.1

            '''         --------------

        End With

    Next

    Set SRF = Cairo.CreateSurface(WorldW, WorldH, ImageSurface)
    Set CC = SRF.CreateContext
    CC.AntiAlias = CAIRO_ANTIALIAS_FAST
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    CC.SetLineWidth 6

    Set GRID = New cSpatialGrid

    GRID.Init WorldW * 1, WorldH * 1, GridSize

    FOLLOWworst
    '    FOLLOWBest

End Sub


Public Sub DRAW()
    Dim I         As Long
    Dim J         As Long
    Dim K         As Long
    Dim CA        As Double
    Dim SA        As Double
    Dim ANG       As Double
    Dim DX#, DY#, D#

    Dim X#, Y#


    CC.SetSourceRGB 0.2, 0.2, 0.2: CC.Paint
    '    CC.SetSourceRGB 0.2, 1, 0.2

    If FOLLOW Then
        CC.Save
        CC.ScaleDrawings ZOOM, ZOOM
        CC.TranslateDrawings WorldW * 0.5 * InvZOOM - Agnet(FOLLOW).X, WorldH * 0.5 * InvZOOM - Agnet(FOLLOW).Y
    End If


    ' GRID
    For X = 0 To WorldW Step GridSize
        CC.DrawLine X, 0, X, WorldH, , 1, 5263440
    Next
    For Y = 0 To WorldH Step GridSize
        CC.DrawLine 0, Y, WorldW, Y, , 1, 5263440
    Next
    '---------

    CC.SetLineWidth 2.5           'Per Spalle

    For I = 1 To NA
        With Agnet(I)
            CC.SetSourceRGB .cR, .cG, .cB
            CC.Arc .X, .Y, 3      ' 3.5
            CC.Fill

            ANG = Atan2(.VX, .VY)
            CA = Cos(ANG)
            SA = Sin(ANG)
            
            CC.MoveTo .X - SA * 3.5, .Y + CA * 3.5
            CC.LineTo .X + CA * 5, .Y + SA * 5
            CC.LineTo .X + SA * 3.5, .Y - CA * 3.5
            CC.LineTo .X - SA * 3.5, .Y + CA * 3.5
            CC.Fill

            ' SHOULDERS-----
            DX = .ShoulderX * 0.87 + 0.13 * CA
            DY = .ShoulderY * 0.87 + 0.13 * SA
            D = 1# / Sqr(DX * DX + DY * DY)
            DX = DX * D: DY = DY * D
            CC.MoveTo .X - DY * 4, .Y + DX * 4
            CC.LineTo .X + DY * 4, .Y - DX * 4
            CC.Stroke
            .ShoulderX = DX
            .ShoulderY = DY
            '-----------------


            ' TARGET
            CC.Arc .TX, .TY, 0.1
            CC.Stroke


            '-------- TRAILS
            ''            For J = 1 To Ntrail
            ''                K = (TRAILS + J) Mod (Ntrail + 1)
            ''                'CC.Arc .SciaX(J), .SciaY(J), 0.7
            ''                'CC.Stroke
            ''                CC.LineTo .SciaX(K), .SciaY(K)
            ''            Next
            ''            CC.Stroke
            '------------



        End With
    Next

    ''X = Agnet(1).X + Agnet(1).VX * VMUL
    ''Y = Agnet(1).Y + Agnet(1).VY * VMUL
    ''CC.Arc X, Y, ControlDist
    ''CC.Stroke


    If FOLLOW Then CC.Restore

    fMain.Picture = SRF.Picture

End Sub

Public Sub MOVE()
    Dim I         As Long
    Dim DX#, DY#
    Dim DX2#, DY2#
    Dim D#
    Dim D2#
    Dim ANG#

    For I = 1 To NA
        With Agnet(I)

            ''--------- Limitazione sterzo
            ''            DX = .VX: DY = .VY
            ''            D = (DX * DX + DY * DY)
            ''            If D Then D = 1 / Sqr(D)
            ''            DX = DX * D
            ''            DY = DY * D
            ''            DX2 = .VXchange: DY2 = .VYchange
            ''            D2 = (DX2 * DX2 + DY2 * DY2)
            ''            If D2 Then D2 = 1 / Sqr(D2)
            ''            DX2 = DX2 * D2
            ''            DY2 = DY2 * D2
            ''
            ''            D = -(DX * DX2 + DY * DY2)
            ''            If D > 0 Then D = -D
            ''            D = 1 + D
            ''
            ''            If D < 0 Then D = 0
            ''            If D > 1 Then D = 1
            ''
            ''            If I = 1 Then Debug.Print D
            ''---------------------
            'D = 1

            .VX = .VX + .VXchange    '* D
            .VY = .VY + .VYchange    '* D

            DX = .TX - .X
            DY = .TY - .Y

            D = (DX * DX + DY * DY)
            If D < 25 Then        ' TARGET REACED
                .NReachedTargets = .NReachedTargets + 1&

                Select Case MODE
                Case 0
                    ' .TX = (0.1 + 0.8 * Rnd) * WorldW
                    ' .TY = (0.1 + 0.8 * Rnd) * WorldH
                    If .TX < WorldW * 0.5 Then
                        .TX = WorldW * 0.9
                    Else
                        .TX = WorldW * 0.1

                    End If


                Case 1
                    '                '-----
                    If .TX < WorldW * 0.5 Then
                        .TX = WorldW * 0.9
                        If I Mod 2 = 0 Then
                            .TY = (0.1 + Rnd * 0.4) * WorldH
                        Else
                            .TY = (0.5 + Rnd * 0.4) * WorldH
                        End If

                    Else
                        .TX = WorldW * 0.1
                        If I Mod 2 <> 0 Then

                            .TY = (0.1 + Rnd * 0.4) * WorldH
                        Else
                            .TY = (0.5 + Rnd * 0.4) * WorldH
                        End If

                    End If
                    '                '-----

                Case 2
                    'CIRCLE
                    .TX = WorldW * 0.5 - (.TX - WorldW * 0.5)
                    .TY = WorldH * 0.5 - (.TY - WorldH * 0.5)


                    D = I / NA * PI2
                    DX = Cos(D): DY = Sin(D)
                    If .NReachedTargets Mod 2 = 0 Then DX = -DX: DY = -DY: D = D - PI: If D < 0 Then D = D + PI2
                    .TX = WorldW * 0.5 + WorldW * 0.4 * DX
                    .TY = WorldH * 0.5 + WorldH * 0.4 * DY

                    '.cR = 0.6
                    '.cG = D / PI2
                    '.cB = 1 - D / PI2
                    .cR = 0.6 + 0.4 * Cos(D * 2 + PI)
                    .cG = 0.6 + 0.4 * Cos(D * 3)
                    .cB = 0.6 + 0.4 * Cos(D * 1 + 2)



                Case 3            'Horiz / VERT
                    If I Mod 2 = 0 Then
                        If .TX < WorldW * 0.5 Then
                            .TX = WorldW * 0.9
                        Else
                            .TX = WorldW * 0.1
                        End If

                        If .TY < WorldH * 0.2 Or .TY > WorldH * 0.8 Then .TY = (0.2 + Rnd * 0.6) * WorldH
                    Else
                        If .TY < WorldH * 0.5 Then
                            .TY = WorldH * 0.9
                        Else
                            .TY = WorldH * 0.1
                        End If
                        If .TX < WorldW * 0.2 Or .TX > WorldW * 0.8 Then .TX = (0.2 + Rnd * 0.6) * WorldW

                    End If


                Case 4

                    D = I / NA * PI2
                    DX = Cos(D): DY = Sin(D)
                    ANG = D


                    If .NReachedTargets Mod 2 = 0 Then DX = -DX: DY = -DY: D = D - PI: If D < 0 Then D = D + PI2

                    DX = DX + Sin(ANG * 2) * 0.1
                    DY = DY + Cos(ANG * 4) * 0.1



                    .TX = WorldW * 0.5 + WorldW * 0.37 * DX
                    .TY = WorldH * 0.5 + WorldH * 0.37 * DY





                    .cR = 0.6 + 0.4 * Cos(D * 2 + PI)
                    .cG = 0.6 + 0.4 * Cos(D * 3)
                    .cB = 0.6 + 0.4 * Cos(D * 1 + 2)




                End Select

            Else
                DX = DX - .VX * VMUL * 1.5
                DY = DY - .VY * VMUL * 1.5

                D = DX * DX + DY * DY

                D = 1# / Sqr(D)
                DX = DX * D
                DY = DY * D

                .VX = .VX + DX * VelChangeGlobStrngth * 0.11
                .VY = .VY + DY * VelChangeGlobStrngth * 0.11

                D = .VX * .VX + .VY * .VY
                If D > MaxV2 Then
                    D = MaxV / Sqr(D)
                    .VX = .VX * D
                    .VY = .VY * D
                End If

            End If

            .X = .X + .VX
            .Y = .Y + .VY

            If .X < 0# Then .X = 0#
            If .Y < 0# Then .Y = 0#
            If .X > WorldW Then .X = WorldW
            If .Y > WorldH Then .Y = WorldH

            .VXchange = 0
            .VYchange = 0

            .V = MaxV2 * 0.5 + .VX * .VX + .VY * .VY


        End With
    Next

    RVO
End Sub

Private Sub RVO()
    Dim pair      As Long
    Dim I         As Long
    Dim J         As Long

    Dim P1()      As Long
    Dim P2()      As Long
    Dim DX()      As Double
    Dim DY()      As Double
    Dim D()       As Double
    Dim Npairs    As Long

    Dim checkX    As Double
    Dim checkY    As Double
    Dim ddX       As Double
    Dim ddY       As Double
    Dim dd        As Double
    Dim Ki        As Double
    Dim Kj        As Double

    Dim MaxVelBasedStrength As Double
    Dim Avoid As Double

    MaxVelBasedStrength = MaxV * 0.75

    With GRID
        .ResetPoints
        For I = 1 To NA
            .InsertPoint Agnet(I).X, Agnet(I).Y
        Next
        .GetPairsWDist P1(), P2(), DX(), DY(), D(), Npairs


        For pair = 1 To Npairs

            I = P1(pair)
            J = P2(pair)

            checkX = Agnet(I).X + (Agnet(I).VX - Agnet(J).VX) * VMUL
            checkY = Agnet(I).Y + (Agnet(I).VY - Agnet(J).VY) * VMUL

            ddX = Agnet(J).X - checkX
            ddY = Agnet(J).Y - checkY

            dd = ddX * ddX + ddY * ddY
            If dd < ControlDist2 Then

                dd = 1# - Sqr(dd) * invControlDist
                ddX = ddX * dd
                ddY = ddY * dd

                ' More changes to Faster
                Ki = Agnet(I).V / (Agnet(I).V + Agnet(J).V)
                Kj = 1 - Ki

                Avoid = Agnet(I).Avoidance

                Agnet(I).VXchange = Agnet(I).VXchange - ddX * Ki * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid
                Agnet(I).VYchange = Agnet(I).VYchange - ddY * Ki * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid

                Avoid = Agnet(J).Avoidance

                Agnet(J).VXchange = Agnet(J).VXchange + ddX * Kj * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid
                Agnet(J).VYchange = Agnet(J).VYchange + ddY * Kj * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid

            End If

        Next
    End With

    If (CNT And 7&) = 0 Then fMain.Caption = "Click to change Mode.   Mode: " & MODE & "   Pairs: " & Npairs & "   Right click to change Camera follow"

End Sub

Public Sub MAINLOOP()
    doloop = True


    Do

        MOVE
        'DRAW
        If (CNT And 1&) = 0& Then DRAW

        CNT = CNT + 1
        DoEvents

        If (CNT And 3&) = 0 Then
            updateScias
        End If


    Loop While doloop

End Sub

Private Sub updateScias()
    Dim I         As Long
    For I = 1 To NA
        Agnet(I).SciaX(TRAILS) = Agnet(I).X
        Agnet(I).SciaY(TRAILS) = Agnet(I).Y
    Next

    TRAILS = (TRAILS + 1) Mod (Ntrail + 1)

End Sub

Public Sub FOLLOWworst()
    Dim R         As Double
    Dim I         As Long
    R = 99
    For I = 1 To NA
        If Agnet(I).Avoidance < R Then
            R = Agnet(I).Avoidance: FOLLOW = I
        End If
    Next
End Sub
Public Sub FOLLOWbest()
    Dim R         As Double
    Dim I         As Long
    R = 0
    For I = 1 To NA
        If Agnet(I).Avoidance > R Then
            R = Agnet(I).Avoidance: FOLLOW = I
        End If
    Next
End Sub

