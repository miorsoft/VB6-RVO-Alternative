Attribute VB_Name = "mRVO"
Option Explicit


Public Type tAgent
    X             As Double
    Y             As Double
    VX            As Double
    VY            As Double
    nVX           As Double
    nVY           As Double

    TX            As Double
    TY            As Double
    VXchange      As Double
    VYchange      As Double

    maxV          As Double
    maxV2         As Double


    cR            As Double
    cG            As Double
    cB            As Double

    V             As Double

    SciaX()       As Double
    SciaY()       As Double

    ShoulderX     As Double
    ShoulderY     As Double

    Avoidance     As Double

    NReachedTargets As Long

    Walked        As Double

End Type

Private TRAILS    As Long

Public Const Ntrail As Long = 0   '1    ' 6    '4

Public Agent()    As tAgent
Public NA         As Long

Public WorldW     As Double
Public WorldH     As Double

Public ScreenW    As Double
Public ScreenH    As Double


Public SRF        As cCairoSurface
Public CC         As cCairoContext

Private Const GlobMaxSpeed As Double = 1  '0.52
Private Const GlobMaxSpeed2 As Double = GlobMaxSpeed * GlobMaxSpeed

Public GRID      As cSpatialGrid

Public doloop     As Boolean

Private CNT       As Long

Private Const ControlDist As Double = 14 '16 '15 ' 14
Private Const ControlDist2 As Double = ControlDist * ControlDist
Private Const invControlDist As Double = 1 / ControlDist

Public Const VMUL As Double = 17 '15 '16 ' 17

Public Const GridSize As Double = 2 * GlobMaxSpeed * VMUL + ControlDist    '40    '35 '25

Private Const VelChangeGlobStrngth As Double = 0.2 ' 0.2 '0.2    'con MaxV

Public MODE       As Long
Public FOLLOW     As Long

Private Const Zoom As Double = 1.3    '1.4
Private Const InvZOOM As Double = 1 / Zoom

Private Const minAvoid As Double = 0.5 '25 '0.2
Private Const OneMinusMinAvoid As Double = 1 - minAvoid

Private TargARR() As tVec3
Private NTar As Long
Private AgTar() As Long



Public Sub Init_RVO(NAG As Long)
    Dim I         As Long
    Dim ANG#




    NA = NAG
    ReDim Agent(NA)
    For I = 1 To NA

        With Agent(I)
            .X = (0.1 + 0.8 * Rnd) * WorldW
            .Y = (0.1 + 0.8 * Rnd) * WorldH
            .TX = (0.1 + 0.8 * Rnd) * WorldW
            .TY = (0.1 + 0.8 * Rnd) * WorldH


            .maxV = GlobMaxSpeed * (0.55 + 0.45 * Rnd)
            If Rnd < 0.01 Then .maxV = 0.2 * GlobMaxSpeed
            
            .maxV2 = .maxV * .maxV

            .Avoidance = minAvoid + OneMinusMinAvoid * Rnd

            ReDim .SciaX(Ntrail)
            ReDim .SciaY(Ntrail)

            .cR = 0

            .cG = (.Avoidance - minAvoid) / (OneMinusMinAvoid)    ': .cB = 0.3
            .cR = 1 - .cG
            .cB = (.maxV / GlobMaxSpeed - 0.55) / 0.45

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

    GRID.INIT WorldW * 1, WorldH * 1, GridSize

    FOLLOWworst
    '    FOLLOWBest

End Sub


Public Sub DRAW()
    Dim I         As Long
    Dim J         As Long
    Dim K         As Long
    Dim ANG       As Double
    Dim DX#, DY#, D#

    Dim X#, Y#


    CC.SetSourceRGB 0.2, 0.2, 0.2: CC.Paint


    If FOLLOW Then
        CC.Save
        CC.ScaleDrawings Zoom, Zoom
        CC.TranslateDrawings WorldW * 0.5 * InvZOOM - Agent(FOLLOW).X, WorldH * 0.5 * InvZOOM - Agent(FOLLOW).Y

        CC.SetSourceRGBA 0.5, 0.5, 0.5, 0.25

        CC.Arc Agent(FOLLOW).X, Agent(FOLLOW).Y, 10
        CC.Stroke

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
        With Agent(I)
            CC.SetSourceRGB .cR, .cG, .cB
            CC.Arc .X, .Y, 3      ' 3.5
            CC.Fill

            ANG = Atan2(.VX, .VY)
            .nVX = Cos(ANG)
            .nVY = Sin(ANG)

            CC.MoveTo .X - .nVY * 3.5, .Y + .nVX * 3.5
            CC.LineTo .X + .nVX * 5#, .Y + .nVY * 5#
            CC.LineTo .X + .nVY * 3.5, .Y - .nVX * 3.5
            CC.LineTo .X - .nVY * 3.5, .Y + .nVX * 3.5
            CC.Fill

            ' SHOULDERS-----
            DX = .ShoulderX * 0.87 + 0.13 * .nVX
            DY = .ShoulderY * 0.87 + 0.13 * .nVY
            D = 1# / Sqr(DX * DX + DY * DY): DX = DX * D: DY = DY * D
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
            For J = 1 To Ntrail
                K = (TRAILS + J) Mod (Ntrail + 1)
                'CC.Arc .SciaX(J), .SciaY(J), 0.7
                'CC.Stroke
                CC.LineTo .SciaX(K), .SciaY(K)
            Next
            CC.Stroke
            '------------



        End With
    Next

    ''X = Agent(1).X + Agent(1).VX * VMUL
    ''Y = Agent(1).Y + Agent(1).VY * VMUL
    ''CC.Arc X, Y, ControlDist
    ''CC.Stroke


    If FOLLOW Then CC.Restore

    fMain.Picture = SRF.Picture


    draw_CAMERA

End Sub
Private Sub CalShoulders()
    Dim I         As Long
    Dim D#, DX#, DY#
    Dim ANG#
    For I = 1 To NA
        With Agent(I)
            ' SHOULDERS-----
            '            ANG = Atan2(.VX, .VY)
            '            .nVX = Cos(ANG)
            '            .nVY = Sin(ANG)

            D = Sqr(.VX * .VX + .VY * .VY)
            .Walked = .Walked + D
            If D Then
                D = 1# / D
                .nVX = .VX * D: .nVY = .VY * D
                DX = .ShoulderX * 0.87 + 0.13 * .nVX
                DY = .ShoulderY * 0.87 + 0.13 * .nVY
                D = 1# / Sqr(DX * DX + DY * DY): DX = DX * D: DY = DY * D
                .ShoulderX = DX
                .ShoulderY = DY

            End If

            '-----------------
        End With
    Next

End Sub
Public Sub draw_CAMERA()
    Dim I         As Long
    Dim L         As tCapsule
    Dim Vis       As Boolean
    Dim R1 As tVec3, R2 As tVec3
    Dim D#
    Dim MUL#
    Dim Kcam      As Double
    Dim DX#, DY#, cX#, cY#

    If FOLLOW Then
        'CAM.lookat = vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)
        '        CAM.SetPositionAndLookAt vec3(WorldW * 0.5, -100, WorldH * 0.5), vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)

        cX = CAM.Position.X
        cY = CAM.Position.Z
        DX = Agent(FOLLOW).X - cX
        DY = Agent(FOLLOW).Y - cY
        If DX * DX + DY * DY > 21100 Then

            cX = cX * 0.992 + 0.008 * Agent(FOLLOW).X
            cY = cY * 0.992 + 0.008 * Agent(FOLLOW).Y

        End If
        CAM.SetPositionAndLookAt vec3(cX, -100, cY), vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)

    End If







    CalShoulders
    Ncap = 0
    For I = 1 To NA
        BUILDH I
    Next

    If Ncap Then QuickSortCapsules CAP(), 1, Ncap


    CC.SetSourceRGB 0.12, 0.2, 0.12: CC.Paint


'--------- GIRD
CC.SetLineWidth 1.5
CC.SetSourceRGB 0.25, 0.25, 0.25
'....................................
For DX = 0 To WorldW Step GridSize
CAM.LineToScreen vec3(DX, 0, 0), vec3(DX, 0, WorldH), R1, R2, Vis
If Vis Then
            CC.MoveTo R1.X, R1.Y
            CC.LineTo R2.X, R2.Y
            CC.Stroke
            End If
Next
For DY = 0 To WorldH Step GridSize
CAM.LineToScreen vec3(0, 0, DY), vec3(WorldW, 0, DY), R1, R2, Vis
If Vis Then
            CC.MoveTo R1.X, R1.Y
            CC.LineTo R2.X, R2.Y
            CC.Stroke
            End If
Next
'-----------------------------



    Kcam = ScreenW * 0.00072



    ' BODYS
    For I = 1 To Ncap
        With CAP(I)
            If .idxA > 0 Then
                CC.SetSourceRGB Agent(.idxA).cR, Agent(.idxA).cG, Agent(.idxA).cB
                MUL = 400
            Else                  'shadow

                CC.SetSourceRGB 0.12, 0.12, 0.16
                MUL = 800
            End If

            CC.SetLineWidth (.sP1.Z + .sP2.Z) * MUL * .Size * Kcam
            CC.MoveTo .sP1.X, .sP1.Y
            CC.LineTo .sP2.X, .sP2.Y
            CC.Stroke
        End With
    Next


    fMain.Picture = SRF.Picture

End Sub

Public Sub MOVE()
    Dim I         As Long
    Dim DX#, DY#
    Dim Dx2#, Dy2#
    Dim D#
    Dim D2#
    Dim ANG#

    For I = 1 To NA
        With Agent(I)

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




                Case 2            'Horiz / VERT
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

                Case 3
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

                Case 4

                    D = I / NA * PI2
                    DX = Cos(D): DY = Sin(D)
                    ANG = D

                    If .NReachedTargets Mod 2 = 0 Then DX = -DX: DY = -DY: D = D - PI: If D < 0 Then D = D + PI2

                    DX = DX + Sin(ANG * 6) * 0.1
                    DY = DY + Cos(ANG * 6) * 0.1

                    .TX = WorldW * 0.5 + WorldW * 0.4 * DX
                    .TY = WorldH * 0.5 + WorldH * 0.4 * DY


                    .cR = 0.6 + 0.4 * Cos(D * 2 + PI)
                    .cG = 0.6 + 0.4 * Cos(D * 3)
                    .cB = 0.6 + 0.4 * Cos(D * 1 + 2)

                Case 5
                    .TX = WorldW * Rnd
                    .TY = WorldH * Rnd

                    .cG = (.Avoidance - minAvoid) / (OneMinusMinAvoid)    ': .cB = 0.3
                    .cR = 1 - .cG
                    .cB = (.maxV / GlobMaxSpeed - 0.55) / 0.45


Case 6
AgTar(I) = (AgTar(I) + 1) Mod (NTar + 1)
.TX = TargARR(AgTar(I)).X + Cos(I / NA * PI2) * 53
.TY = TargARR(AgTar(I)).Y + Sin(I / NA * PI2) * 53



                End Select

            Else
                DX = DX - .VX * VMUL * 1.5
                DY = DY - .VY * VMUL * 1.5

                D = DX * DX + DY * DY

                D = 1# / Sqr(D)
                DX = DX * D
                DY = DY * D

'                .VX = .VX + DX * VelChangeGlobStrngth * 0.125    ' 0.11
'                .VY = .VY + DY * VelChangeGlobStrngth * 0.125    ' 0.11

                .VX = .VX + DX * VelChangeGlobStrngth * 0.125 * .maxV
                .VY = .VY + DY * VelChangeGlobStrngth * 0.125 * .maxV

                D = .VX * .VX + .VY * .VY
                If D > .maxV2 Then
                    D = .maxV / Sqr(D)
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

            .VXchange = .VXchange * 0.5 '0
            .VYchange = .VYchange * 0.5
            

            .V = GlobMaxSpeed2 * 0.5 + .VX * .VX + .VY * .VY


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

    Dim CheckPosX    As Double
    Dim CheckPosY    As Double
    Dim ddX       As Double
    Dim ddY       As Double
    Dim NdX       As Double
    Dim NdY       As Double

    Dim dd        As Double
    Dim Ki        As Double
    Dim Kj        As Double
    Dim JIdx      As Double
    Dim JIdy      As Double
    Dim JID       As Double

    Dim DOT       As Double

    Dim MaxVelBasedStrength As Double
    Dim Avoid     As Double
    Dim SqrDD     As Double


    MaxVelBasedStrength = GlobMaxSpeed * 1.125

    With GRID
        .ResetPoints
        For I = 1 To NA
            .InsertPoint Agent(I).X, Agent(I).Y
        Next
        .GetPairsWDist P1(), P2(), DX(), DY(), D(), Npairs


        For pair = 1 To Npairs

            I = P1(pair)
            J = P2(pair)

            CheckPosX = Agent(I).X + (Agent(I).VX - Agent(J).VX) * VMUL
            CheckPosY = Agent(I).Y + (Agent(I).VY - Agent(J).VY) * VMUL

            ddX = Agent(J).X - CheckPosX
            ddY = Agent(J).Y - CheckPosY

            dd = ddX * ddX + ddY * ddY
            If dd < ControlDist2 Then
                
                SqrDD = Sqr(dd)
                
                'DD inverse proportional to DIST (Range 0,1)
                dd = 1# - SqrDD * invControlDist
                ' 'smoothstep
                dd = dd * dd * (3# - 2# * dd)

'Normalize ddx,ddy and multiply by DD
SqrDD = 1# / SqrDD
                ddX = ddX * SqrDD * dd
                ddY = ddY * SqrDD * dd

                ' More changes to Faster
                Ki = Agent(I).V / (Agent(I).V + Agent(J).V)
                Kj = 1 - Ki
                '

                'Compute Normalized DX DY of Agents I,J
                JIdx = DX(pair)
                JIdy = DY(pair)
                JID = 1# / Sqr(D(pair))
                JIdx = JIdx * JID
                JIdy = JIdy * JID


                DOT = Agent(I).ShoulderX * JIdx + Agent(I).ShoulderY * JIdy
                
                Avoid = Agent(I).Avoidance * (0.7 + DOT * 0.3) 'Mora Avoid if orher agent is in front
                DOT = 0.5 + DOT * 0.5
                DOT = DOT * DOT * 0.005 * Agent(I).maxV
                Agent(I).VXchange = Agent(I).VXchange - ddX * Ki * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid
                Agent(I).VYchange = Agent(I).VYchange - ddY * Ki * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid
                'It's useful to take some Velocity from orher Agent
                Agent(I).VXchange = Agent(I).VXchange + Agent(J).VX * DOT
                Agent(I).VYchange = Agent(I).VYchange + Agent(J).VY * DOT

                DOT = Agent(J).ShoulderX * -JIdx + Agent(J).ShoulderY * -JIdy
                DOT = 0.5 + DOT * 0.5
                Avoid = Agent(J).Avoidance * (0.7 + DOT * 0.3)
                DOT = 0.5 + DOT * 0.5
                DOT = DOT * DOT * 0.005 * Agent(J).maxV
                Agent(J).VXchange = Agent(J).VXchange + ddX * Kj * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid
                Agent(J).VYchange = Agent(J).VYchange + ddY * Kj * VelChangeGlobStrngth * MaxVelBasedStrength * Avoid
                Agent(J).VXchange = Agent(J).VXchange + Agent(I).VX * DOT
                Agent(J).VYchange = Agent(J).VYchange + Agent(I).VY * DOT

            End If


            '------------------------------------
            'Separate
            If D(pair) < 25 Then
                dd = Sqr(D(pair))
                ddX = DX(pair) / dd
                ddY = DY(pair) / dd
                ddX = ddX * (5 - dd)
                ddY = ddY * (5 - dd)
                Agent(I).X = Agent(I).X - ddX
                Agent(I).Y = Agent(I).Y - ddY
                Agent(J).X = Agent(J).X + ddX
                Agent(J).Y = Agent(J).Y + ddY
            End If


        Next
    End With

    If (CNT And 7&) = 0 Then fMain.Caption = "Click to change Mode.   Mode: " & MODE & "   Pairs: " & Npairs & "   Right click to change Camera follow"

End Sub

Public Sub MAINLOOP()
    doloop = True


    Do

        MOVE

        'If (CNT And 1&) = 0& Then DRAW
        If (CNT And 1&) = 0& Then draw_CAMERA



        If (CNT And 3&) = 0 Then
            updateTRAILS
        End If

        CNT = CNT + 1
        DoEvents

    Loop While doloop

End Sub

Private Sub updateTRAILS()
    Dim I         As Long
    For I = 1 To NA
        Agent(I).SciaX(TRAILS) = Agent(I).X
        Agent(I).SciaY(TRAILS) = Agent(I).Y
    Next

    TRAILS = (TRAILS + 1) Mod (Ntrail + 1)

End Sub

Public Sub FOLLOWworst()
    Dim R         As Double
    Dim I         As Long
    R = 99
    For I = 1 To NA
        If Agent(I).Avoidance < R Then
            R = Agent(I).Avoidance: FOLLOW = I
        End If
    Next
End Sub
Public Sub FOLLOWbest()
    Dim R         As Double
    Dim I         As Long
    R = 0
    For I = 1 To NA
        If Agent(I).Avoidance > R Then
            R = Agent(I).Avoidance: FOLLOW = I
        End If
    Next
End Sub



Private Sub BUILDH(I As Long)
    Dim L         As tCapsule

    Dim X#, Y#
    Dim nX#, nY#
    Dim DX#, DY#: Dim A#
    Dim W#: Dim H#


    Dim Ank       As tVec3
    Dim Feet      As tVec3
    Dim Knee      As tVec3
    Dim tmp1      As tVec3
    Dim tmp2      As tVec3

    X = Agent(I).X
    Y = Agent(I).Y
    DX = Agent(I).ShoulderX
    DY = Agent(I).ShoulderY
    W = Agent(I).Walked

    nX = Agent(I).nVX
    nY = Agent(I).nVY
    With L
        .Size = 1.5
        .P1.X = X
        .P1.Z = Y
        .P2 = L.P1
        .P1.Y = -4.5
        .P2.Y = -7.5
        ADDlineToScreen L, I

        '        .P1.Y = -7: .P2.Y = -7
        '        .P1.X = X - DX * 4
        '        .P1.Z = Y - DY * 4
        '        .P2.X = X + DX * 4
        '        .P2.Z = Y + DY * 4
        '        ADDlineToScreen L, I



        '.P1 = vec3(0, -7, 3)
        '.P2 = vec3(0, -7, -3)
        'transform .P1, .P2, DX, DY, X, Y:        ADDlineToScreen L, I
        .Size = 1
        '------------------------------------------------------------------------

        'LEGS
        ''        A = W * 0.58    '0.5
        ''        .P1 = vec3(0, -4.5, 1.25)
        ''        H = Sin(A) * 2.5: If H > 0 Then H = 0
        ''        .P2 = vec3(Cos(A) * 3, H, 1.25)
        ''        transform .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I
        ''
        ''        A = A + PI
        ''        .P1 = vec3(0, -4.5, -1.25)
        ''        H = Sin(A) * 2.5: If H > 0 Then H = 0
        ''        .P2 = vec3(Cos(A) * 3, H, -1.25)
        ''        transform .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I
        ''



        A = W * 0.65              '0.5
        .P1 = vec3(0, -4.5, 1.25)
        H = Sin(A) * 1.5: If H > 0 Then H = 0
        .P2 = vec3(Cos(A) * 3, H, 1.25)
        SOLVEik .P1, .P2, 2.5, 2, Knee, tmp1
        tmp1 = .P1
        tmp2 = .P2
        Knee.Z = .P1.Z
        .P2 = Knee
        transform .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I
        .P1 = Knee
        .P2 = tmp2
        transform .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I


        A = A + PI
        .P1 = vec3(0, -4.5, -1.25)
        H = Sin(A) * 1.5: If H > 0 Then H = 0
        .P2 = vec3(Cos(A) * 3, H, -1.25)
        SOLVEik .P1, .P2, 2.5, 2, Knee, tmp1
        tmp1 = .P1
        tmp2 = .P2
        Knee.Z = .P1.Z
        .P2 = Knee
        transform .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I
        .P1 = Knee
        .P2 = tmp2
        transform .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I
        '------------------------------------------------------------------------

        'ARMS
        .P1 = vec3(0, -8, 1.5)
        .P2 = vec3(Cos(A) + 0.5, -4 - Cos(A), 3)
        transform .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I

        .P1 = vec3(0, -8, -1.5)
        .P2 = vec3(Cos(A + PI) + 0.5, -4 - Cos(A + PI), -3)
        transform .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I

        'HEAD
        .P1 = vec3(0.2, -9.5, 0)
        .P2 = vec3(0.4, -10, 0)
        transform .P1, .P2, nX, nY, X, Y: L.Size = 1.75: ADDlineToScreen L, I





        'TARGET
        .P1.X = Agent(I).TX
        .P1.Z = Agent(I).TY
        .P1.Y = 0
        .P2 = L.P1
        .P2.Y = -0.1
        .Size = 1.5
        ADDlineToScreen L, I
    End With
End Sub

Private Sub RotateXZ(p As tVec3, DX#, DY#)
    Dim c         As tVec3: c = p
    p.X = c.X * DX + c.Z * -DY
    p.Z = c.X * DY + c.Z * DX
End Sub
Private Sub AddXZ(p As tVec3, X#, Y#)
    p.X = p.X + X
    p.Z = p.Z + Y
End Sub

Private Sub transform(P1 As tVec3, P2 As tVec3, DX#, DY#, X#, Y#)
'    RotateXZ P1, DX, DY
'    RotateXZ P2, DX, DY
'    AddXZ P1, X, Y
'    AddXZ P2, X, Y
  Dim c         As tVec3: c = P1
    P1.X = c.X * DX + c.Z * -DY + X
    P1.Z = c.X * DY + c.Z * DX + Y

c = P2
    P2.X = c.X * DX + c.Z * -DY + X
    P2.Z = c.X * DY + c.Z * DX + Y

End Sub


Public Sub INIT_Targets()
    Dim I         As Long
    Dim FirstTarg As Boolean

    If NTar = 0 Then
        ReDim AgTar(NA)
        FirstTarg = True

    End If

    NTar = 3 + 2
    ReDim TargARR(NTar)



    For I = 0 To NTar
        TargARR(I).X = 200 + Rnd * (WorldW - 400)
        TargARR(I).Y = 200 + Rnd * (WorldH - 400)
    Next

    If FirstTarg Then
        For I = 1 To NA
            Agent(I).TX = TargARR(0).X + Cos(I / NA * PI2) * 53
            Agent(I).TY = TargARR(0).Y + Sin(I / NA * PI2) * 53
        Next
    End If


End Sub
