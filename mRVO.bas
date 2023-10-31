Attribute VB_Name = "mRVO"
Option Explicit


Public Type tAgent
    X             As Double
    Y             As Double
    vX            As Double
    vY            As Double
    nVX           As Double       'Normalized Vel
    nVY           As Double

    TX            As Double       'Target Position
    TY            As Double
    VXchange      As Double
    VYchange      As Double

    maxV          As Double
    maxV2         As Double       ' MaxV*MaxV

    cR            As Double
    cG            As Double
    cB            As Double

    v             As Double

    SciaX()       As Double
    SciaY()       As Double

    ShoulderX     As Double       'Shoulder/Chest direction
    ShoulderY     As Double

    Avoidance     As Double

    NReachedTargets As Long

    Walked        As Double       'Walked Distance

    'VTTS As Double


    fMinD         As Double
    fTargTo       As Long
    fAgentFrom    As Long

End Type




Public Type tPairInfo
    a             As Long
    b             As Long
    chkX          As Double
    chkY          As Double
    chkX2         As Double
    chkY2         As Double

    aVX           As Double
    aVY           As Double
    bVX           As Double
    bVY           As Double

    aVXch         As Double
    aVYch         As Double
    bVXch         As Double
    bVYch         As Double

    Enabled       As Long
    GTC           As Double

End Type

Public PairInfo() As tPairInfo




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

Private Const GlobMaxSpeed As Double = 1
Private Const GlobMaxSpeed2 As Double = GlobMaxSpeed * GlobMaxSpeed
Private Const INVGlobMaxSpeed As Double = 1 / GlobMaxSpeed

Public GRID       As cSpatialGrid

Public doLOOP     As Boolean
Public doFRAMES   As Boolean
Private Frame     As Long

Private CNT       As Long

Public Const ControlDist As Double = 28    ' 28.5 '19   ' 17    '17
Private Const ControlDist2 As Double = ControlDist * ControlDist
Private Const invControlDist As Double = 1 / ControlDist

Public Const VMUL As Double = 14  ' 27  '20  '21    '17

Public Const GridSize As Double = 2 * GlobMaxSpeed * VMUL + ControlDist    '40    '35 '25

Private Const toTargetStrengthK As Double = 0.018    ' 0.02    '0.2*0.125


Private Const AgentsMinDist As Double = 5
Private Const AgentsMinDist2 As Double = AgentsMinDist * AgentsMinDist

'Private Const ControlDistPERC As Double = 1 - 0.5 * AgentsMinDist / ControlDist

Public MODE       As Long
Public FOLLOW     As Long
Public FollowMode As Long
Public Const NFollowModes As Long = 10
Public FollowDesc As String
Public MAXTR      As Long
'Private Const Zoom As Double = 1.3    '1.4
'Private Const InvZOOM As Double = 1 / Zoom

Private Const KglobAvoidance As Double = 1
Private Const minAvoid As Double = 0.5    '25 '0.2
Private Const OneMinusMinAvoid As Double = 1 - minAvoid

' FOR MODE 6
Public ArrayOfTargets() As tVec3
Public NTargets   As Long
Private AgentTarget() As Long

Public RunningNotInIDE As Boolean

Public CameraMode As Boolean

Public RVONpairs  As Long

Public Ncollisions As Long


Private Sub SetColor(Avoidance#, MaxVel#, r, G#, b#)

    'G = (Avoidance - minAvoid) / (OneMinusMinAvoid)    ': .cB = 0.3
    'R = 1 - G
    'B = (MaxVel / GlobMaxSpeed - 0.55) / 0.45

    G = (Avoidance / KglobAvoidance - minAvoid) / (OneMinusMinAvoid)    ': .cB = 0.3
    r = 1 - G
    b = (MaxVel / GlobMaxSpeed - 0.5) / 0.5

End Sub

Public Sub Init_RVO(NAG As Long)
    Dim I         As Long
    Dim ANG#




    NA = NAG
    ReDim Agent(NA)
    ReDim FEETs(NA)

    For I = 1 To NA

        With Agent(I)
            .X = (0.1 + 0.8 * Rnd) * WorldW
            .Y = (0.1 + 0.8 * Rnd) * WorldH
            .TX = (0.1 + 0.8 * Rnd) * WorldW
            .TY = (0.1 + 0.8 * Rnd) * WorldH


            .maxV = GlobMaxSpeed * (0.55 + 0.45 * Rnd)
            '            If Rnd < 0.01 Then .maxV = 0.2 * GlobMaxSpeed

            .maxV2 = .maxV * .maxV

            .Avoidance = (minAvoid + OneMinusMinAvoid * Rnd) * KglobAvoidance

            ReDim .SciaX(Ntrail)
            ReDim .SciaY(Ntrail)



            SetColor .Avoidance, .maxV, .cR, .cG, .cB


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

    '    FOLLOWworstrule
    FOLLOW = FOLLOWbestRule

End Sub


Public Sub DRAW()
    Dim I         As Long
    Dim J         As Long
    Dim K         As Long
    Dim ANG       As Double
    Dim dx#, dy#, d#

    Dim X#, Y#
    Const Zoom    As Double = 1
    Const InvZoom As Double = 1 / Zoom



    CC.SetSourceRGB 0.2, 0.2, 0.2: CC.Paint


    If FOLLOW Then
        CC.Save
        CC.ScaleDrawings Zoom, Zoom
        CC.TranslateDrawings WorldW * 0.5 * InvZoom - Agent(FOLLOW).X, WorldH * 0.5 * InvZoom - Agent(FOLLOW).Y

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

            ANG = Atan2(.vX, .vY)
            .nVX = Cos(ANG)
            .nVY = Sin(ANG)

            CC.MoveTo .X - .nVY * 3.5, .Y + .nVX * 3.5
            CC.LineTo .X + .nVX * 5#, .Y + .nVY * 5#
            CC.LineTo .X + .nVY * 3.5, .Y - .nVX * 3.5
            CC.LineTo .X - .nVY * 3.5, .Y + .nVX * 3.5
            CC.Fill

            ' SHOULDERS-----
            dx = .ShoulderX * 0.87 + 0.13 * .nVX
            dy = .ShoulderY * 0.87 + 0.13 * .nVY
            d = 1# / Sqr(dx * dx + dy * dy): dx = dx * d: dy = dy * d
            CC.MoveTo .X - dy * 4, .Y + dx * 4
            CC.LineTo .X + dy * 4, .Y - dx * 4
            CC.Stroke
            .ShoulderX = dx
            .ShoulderY = dy
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


Private Sub Draw_PAIR()
    Dim X         As Double
    Dim Y         As Double
    Dim Zoom      As Double
    Dim InvZoom   As Double
    Dim I         As Long
    Dim J         As Long


    Zoom = 5
    InvZoom = 1 / Zoom

    CalShoulders                  'as draw_camera


    CC.SetSourceRGB 0.2, 0.2, 0.2: CC.Paint


    CC.Save

    CC.ScaleDrawings Zoom, Zoom
    CC.TranslateDrawings ScreenW * 0.5 * InvZoom - Agent(FOLLOW).X, ScreenH * 0.5 * InvZoom - Agent(FOLLOW).Y

    CC.SetLineWidth 0.1

    For X = 0 To WorldW Step GridSize
        CC.DrawLine X, 0, X, WorldH, , 1, 5263440
    Next
    For Y = 0 To WorldH Step GridSize
        CC.DrawLine 0, Y, WorldW, Y, , 1, 5263440
    Next

    CC.SetLineWidth 1

    For I = 1 To RVONpairs
        If PairInfo(I).a = FOLLOW Or PairInfo(I).b = FOLLOW Then
            If PairInfo(I).Enabled Then

                With PairInfo(I)

                    If .b = FOLLOW Then    'SWAPS
                        J = .b: .b = .a: .a = J
                        X = .bVX: .bVX = .aVX: .aVX = X
                        Y = .bVY: .bVY = .aVY: .aVY = Y
                        X = .bVXch: .bVXch = .aVXch: .aVXch = X
                        Y = .bVYch: .bVYch = .aVYch: .aVYch = Y
                        X = .chkX: .chkX = .chkX2: .chkX2 = X
                        Y = .chkY: .chkY = .chkY2: .chkY2 = Y
                    End If

                    CC.SetSourceRGB Agent(.a).cR, Agent(.a).cG, Agent(.a).cB
                    '                    CC.Arc Agent(.A).X, Agent(.A).Y, AgentsMinDist * 0.5
                    '                    CC.Stroke

                    CC.MoveTo Agent(.a).X, Agent(.a).Y
                    CC.RelLineTo .aVX, .aVY
                    CC.Stroke

                    CC.MoveTo Agent(.a).X + .aVX, Agent(.a).Y + .aVY
                    CC.RelLineTo .aVXch * VMUL * 50, .aVYch * VMUL * 50
                    CC.Stroke
                    '-----------------

                    CC.SetSourceRGB Agent(.b).cR, Agent(.b).cG, Agent(.b).cB
                    '                    CC.Arc Agent(.B).X, Agent(.B).Y, AgentsMinDist * 0.5
                    '                    CC.Stroke

                    CC.MoveTo Agent(.b).X, Agent(.b).Y
                    CC.RelLineTo .bVX, .bVY
                    CC.Stroke

                    CC.MoveTo Agent(.b).X + .bVX, Agent(.b).Y + .bVY
                    CC.RelLineTo .bVXch * VMUL * 50, .bVYch * VMUL * 50
                    CC.Stroke

                    'checkPt
                    CC.SetSourceRGBA 1, 1, 0, 0.05
                    CC.Arc .chkX, .chkY, ControlDist: CC.Fill
                    CC.SetSourceRGB 1, 1, 0
                    CC.Arc .chkX, .chkY, .GTC * ControlDist: CC.Stroke


                    '.................................
                    CC.SetSourceRGB 0.75, 0.75, 0.75
                    ' CC.MoveTo Agent(.A).X, Agent(.A).Y
                    CC.MoveTo Agent(.b).X, Agent(.b).Y
                    CC.LineTo .chkX, .chkY: CC.Stroke




                End With





            End If
        End If

    Next



    CC.SetLineWidth 1
    For J = 1 To NA
        CC.SetSourceRGB Agent(J).cR, Agent(J).cG, Agent(J).cB
        CC.Arc Agent(J).TX, Agent(J).TY, 0.5
        CC.Stroke
    Next
    For J = 1 To NA
        CC.SetSourceRGB Agent(J).cR, Agent(J).cG, Agent(J).cB
        CC.Arc Agent(J).X, Agent(J).Y, AgentsMinDist * 0.5
        CC.Fill
        CC.MoveTo Agent(J).X, Agent(J).Y
        CC.RelLineTo Agent(J).vX * VMUL, Agent(J).vY * VMUL
        CC.Stroke

        '        CC.MoveTo Agent(J).X - Agent(J).ShoulderY * AgentsMinDist * 0.8, Agent(J).Y + Agent(J).ShoulderX * AgentsMinDist * 0.8
        '        CC.RelLineTo Agent(J).ShoulderY * AgentsMinDist * 1.6, -Agent(J).ShoulderX * AgentsMinDist * 1.6
        '        CC.Stroke

        CC.MoveTo Agent(J).X, Agent(J).Y
        CC.RelLineTo Agent(J).ShoulderX * AgentsMinDist * 1, Agent(J).ShoulderY * AgentsMinDist * 1
        CC.Stroke
    Next
    CC.Restore


    OUTINFO

    fMain.Picture = SRF.Picture

End Sub








Private Sub CalShoulders()

    'Compute Shoulder/chest/arm direction by smoothing Normalized Velocity
    Dim I         As Long
    Dim d#, dx#, dy#
    Dim ANG#
    For I = 1 To NA
        With Agent(I)
            ' SHOULDERS-----
            d = Sqr(.vX * .vX + .vY * .vY)
            .v = d

            .Walked = .Walked + d + 0.05
            If d Then
                d = 1# / d
                .nVX = .vX * d: .nVY = .vY * d
                '                '                'DX = .ShoulderX * 0.87 + 0.13 * .nVX
                '                '                'DY = .ShoulderY * 0.87 + 0.13 * .nVY
                '                '                DX = .ShoulderX * 0.85 + 0.15 * .nVX
                '                '                DY = .ShoulderY * 0.85 + 0.15 * .nVY
                '                ' ------ SHOULDERS 1St Mode (works better quando blocco rotazione) [AAA]
                '                                DX = .ShoulderX * 0.86 + 0.14 * .nVX
                '                                DY = .ShoulderY * 0.86 + 0.14 * .nVY
                '                                D = 1# / Sqr(DX * DX + DY * DY): DX = DX * D: DY = DY * D
                '                                .ShoulderX = DX
                '                                .ShoulderY = DY
            End If
            '------------ SHOLDERS 2ND mode
            .ShoulderX = .ShoulderX * 0.7 + 0.3 * .vX * INVGlobMaxSpeed
            .ShoulderY = .ShoulderY * 0.7 + 0.3 * .vY * INVGlobMaxSpeed

            d = .ShoulderX * .ShoulderX + .ShoulderY * .ShoulderY
            If d Then
                d = 1# / Sqr(d)
                .ShoulderX = .ShoulderX * d
                .ShoulderY = .ShoulderY * d
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
    Dim d#
    Const MUL     As Double = 400

    Dim Kcam      As Double
    Dim dx#, dy#, cX#, cY#
    Dim DOT#
    Dim tCamLookAt As tVec3

    If FOLLOW Then

        tCamLookAt = CAM.lookat



        If MODE <> 7 Then

            'CAM.lookat = vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)
            '        CAM.SetPositionAndLookAt vec3(WorldW * 0.5, -100, WorldH * 0.5), vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)
            cX = CAM.Position.X
            cY = CAM.Position.Z
            dx = Agent(FOLLOW).X - cX
            dy = Agent(FOLLOW).Y - cY
            '        If DX * DX + DY * DY > 21100 Then
            '            cX = cX * 0.992 + 0.008 * Agent(FOLLOW).X
            '            cY = cY * 0.992 + 0.008 * Agent(FOLLOW).Y
            '        End If
            If dx * dx + dy * dy > 15000 Then
                cX = cX * 0.992 + 0.008 * Agent(FOLLOW).X
                cY = cY * 0.992 + 0.008 * Agent(FOLLOW).Y
            End If
            '        CAM.SetPositionAndLookAt vec3(cX, -100, cY), vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)
            CAM.SetPositionAndLookAt vec3(cX, -85, cY), _
                                     vec3(tCamLookAt.X * 0.8 + 0.2 * Agent(FOLLOW).X, -5#, tCamLookAt.Z * 0.8 + 0.2 * Agent(FOLLOW).Y)

        Else

            cX = CAM.Position.X
            cY = CAM.Position.Z
            dx = Agent(FOLLOW).X - cX
            dy = Agent(FOLLOW).Y - cY
            '        If DX * DX + DY * DY > 21100 Then
            '            cX = cX * 0.992 + 0.008 * Agent(FOLLOW).X
            '            cY = cY * 0.992 + 0.008 * Agent(FOLLOW).Y
            '        End If
            If dx * dx + dy * dy > 15000 Then
                cX = cX * 0.992 + 0.008 * Agent(FOLLOW).X
                cY = cY * 0.992 + 0.008 * Agent(FOLLOW).Y
            End If
            '        CAM.SetPositionAndLookAt vec3(cX, -100, cY), vec3(Agent(FOLLOW).X, 0, Agent(FOLLOW).Y)
            CAM.SetPositionAndLookAt vec3(cX, -160 * 2, cY), _
                                     vec3(tCamLookAt.X * 0.85 + 0.15 * Agent(FOLLOW).X, 0, tCamLookAt.Z * 0.85 + 0.15 * Agent(FOLLOW).Y)

        End If
    End If







    CalShoulders


    NCapsulesInScreen = 0
    For I = 1 To NA: BUILDHuman I: Next

    If NCapsulesInScreen Then QuickSortCapsules CAPSULES(), 1, NCapsulesInScreen


    '    CC.SetSourceRGB 0.12, 0.2, 0.12: CC.Paint
    CC.SetSourceRGB 0.13, 0.216, 0.13: CC.Paint


    '--------- GIRD
    CAM.FarPlane = 5000
    CC.SetLineWidth 1.5
    CC.SetSourceRGB 0.25, 0.25, 0.25
    '....................................
    For dx = 0 To WorldW Step GridSize
        CAM.LineToScreen vec3(dx, 0, 0), vec3(dx, 0, WorldH), R1, R2, Vis
        If Vis Then
            CC.MoveTo R1.X, R1.Y
            CC.LineTo R2.X, R2.Y
            CC.Stroke
        End If
    Next
    For dy = 0 To WorldH Step GridSize
        CAM.LineToScreen vec3(0, 0, dy), vec3(WorldW, 0, dy), R1, R2, Vis
        If Vis Then
            CC.MoveTo R1.X, R1.Y
            CC.LineTo R2.X, R2.Y
            CC.Stroke
        End If
    Next
    '-----------------------------



    Kcam = ScreenW * 0.00072



    ' BODYS
    For I = 1 To NCapsulesInScreen
        With CAPSULES(I)
            If .AgentIndex > 0 Then
                If (.ScreenP1.Z + .ScreenP2.Z) > 0.0028 Then    'Countour

                    CC.SetLineWidth (.ScreenP1.Z + .ScreenP2.Z) * (MUL * .Size * Kcam + MUL * Kcam * 0.66)
                    CC.SetSourceRGB 0, 0, 0
                    CC.MoveTo .ScreenP1.X, .ScreenP1.Y
                    CC.LineTo .ScreenP2.X, .ScreenP2.Y
                    CC.Stroke
                    '----------------------------

                    ''                    D = (.ScreenP1.Z + .ScreenP2.Z) * MUL * .Size * Kcam
                    ''                    CC.SetLineWidth D
                    ''                    D = (.ScreenP1.Z + .ScreenP2.Z) * MUL * Kcam * 0.2
                    ''                    dx = D * (CAM.Direction.X - CAM.Direction.Z)
                    ''                    DY = D * (CAM.Direction.X + CAM.Direction.Z)
                    ''                    CC.SetSourceRGB Agent(.AgentIndex).cR * 1.5, Agent(.AgentIndex).cG * 1.5, Agent(.AgentIndex).cB * 1.5
                    ''                    CC.MoveTo .ScreenP1.X - dx, .ScreenP1.Y - DY
                    ''                    CC.LineTo .ScreenP2.X - dx, .ScreenP2.Y - DY
                    ''                    CC.Stroke
                    ''
                    ''                    dx = dx * 2: DY = DY * 2
                    ''                    CC.SetSourceRGBA 0, 0, 0, 1
                    ''                    CC.MoveTo .ScreenP1.X + dx, .ScreenP1.Y + DY
                    ''                    CC.LineTo .ScreenP2.X + dx, .ScreenP2.Y + DY
                    ''                    CC.Stroke
                End If
                CC.SetSourceRGB Agent(.AgentIndex).cR, Agent(.AgentIndex).cG, Agent(.AgentIndex).cB
                CC.SetLineWidth (.ScreenP1.Z + .ScreenP2.Z) * MUL * .Size * Kcam
                CC.MoveTo .ScreenP1.X, .ScreenP1.Y
                CC.LineTo .ScreenP2.X, .ScreenP2.Y
                CC.Stroke
            Else                  'SHADOW
                CC.SetSourceRGB 0.12, 0.12, 0.16
                CC.SetLineWidth (.ScreenP1.Z + .ScreenP2.Z) * MUL * .Size * Kcam * 2#    ''* 2 is due to Fake far away for shadows
                CC.MoveTo .ScreenP1.X, .ScreenP1.Y
                CC.LineTo .ScreenP2.X, .ScreenP2.Y
                CC.Stroke
            End If


        End With
    Next

    'MIRINO -----------------------------------------
    With CC
        .Save
        .SetDashes 40 * PI2 * 0.125, 80 * PI2 * 0.125, 80 * PI2 * 0.125
        .SetLineWidth 1
        .SetSourceRGBA 1, 1, 1, 0.15
        .Arc ScreenW * 0.5, ScreenH * 0.5, 80
        .Stroke
        .SetDashes -12.5, 25, 25
        .MoveTo ScreenW * 0.5, ScreenH * 0.5 - 50
        .LineTo ScreenW * 0.5, ScreenH * 0.5 + 50: .Stroke
        .MoveTo ScreenW * 0.5 - 50, ScreenH * 0.5
        .LineTo ScreenW * 0.5 + 50, ScreenH * 0.5: .Stroke
        .Restore

    End With

    OUTINFO

    fMain.Picture = SRF.Picture

End Sub

Public Sub MOVE()
    Dim I         As Long
    Dim dx#, dy#
    Dim Dx2#, Dy2#
    Dim d#
    Dim D2#
    Dim ANG#
    Dim a#, r#

    For I = 1 To NA
        With Agent(I)

            .vX = .vX + .VXchange
            .vY = .vY + .VYchange

            dx = .TX - .X
            dy = .TY - .Y

            d = (dx * dx + dy * dy)
            If d < 4 Then         '25    ' TARGET REACHED


                SetColor .Avoidance, .maxV, .cR, .cG, .cB

                Select Case MODE
                Case 0            '2 directions
                    ' .TX = (0.1 + 0.8 * Rnd) * WorldW
                    ' .TY = (0.1 + 0.8 * Rnd) * WorldH
                    If .TX < WorldW * 0.5 Then
                        .TX = WorldW * 0.9
                    Else
                        .TX = WorldW * 0.1

                    End If

                Case 1            'diagonal
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


                    d = I / NA * PI2
                    dx = Cos(d): dy = Sin(d)
                    If .NReachedTargets Mod 2 = 0 Then dx = -dx: dy = -dy: d = d - PI: If d < 0 Then d = d + PI2
                    .TX = WorldW * 0.5 + WorldW * 0.4 * dx
                    .TY = WorldH * 0.5 + WorldH * 0.4 * dy

                    '.cR = 0.6
                    '.cG = D / PI2
                    '.cB = 1 - D / PI2
                    .cR = 0.6 + 0.4 * Cos(d * 2 + PI)
                    .cG = 0.6 + 0.4 * Cos(d * 3)
                    .cB = 0.6 + 0.4 * Cos(d * 1 + 2)

                    '   Case 4

                    '                    D = I / NA * PI2
                    '                    DX = Cos(D): DY = Sin(D)
                    '                    ANG = D
                    '
                    '                    If .NReachedTargets Mod 2 = 0 Then DX = -DX: DY = -DY: D = D - PI: If D < 0 Then D = D + PI2
                    '
                    '                    DX = DX + Sin(ANG * 6) * 0.1
                    '                    DY = DY + Cos(ANG * 6) * 0.1
                    '
                    '                    .TX = WorldW * 0.5 + WorldW * 0.4 * DX
                    '                    .TY = WorldH * 0.5 + WorldH * 0.4 * DY
                    '
                    '
                    '                    .cR = 0.6 + 0.4 * Cos(D * 2 + PI)
                    '                    .cG = 0.6 + 0.4 * Cos(D * 3)
                    '                    .cB = 0.6 + 0.4 * Cos(D * 1 + 2)
                Case 4            'Horiz / VERT narrow cross

                    If I Mod 2 = 0 Then
                        If .TX < WorldW * 0.5 Then
                            .TX = WorldW * 0.9
                        Else
                            .TX = WorldW * 0.1
                        End If

                        If .TY < WorldH * 0.4 Or .TY > WorldH * 0.6 Then .TY = (0.4 + Rnd * 0.2) * WorldH
                    Else
                        If .TY < WorldH * 0.5 Then
                            .TY = WorldH * 0.9
                        Else
                            .TY = WorldH * 0.1
                        End If
                        If .TX < WorldW * 0.4 Or .TX > WorldW * 0.6 Then .TX = (0.4 + Rnd * 0.2) * WorldW

                    End If
                Case 5            'RANDOM
                    .TX = WorldW * Rnd
                    .TY = WorldH * Rnd


                Case 6            'TARGETS
                    AgentTarget(I) = (AgentTarget(I) + 1) Mod (NTargets + 1)
                    r = I / NA: r = Sqr(r)
                    '            R = 10 + 99 * R
                    r = 4 + r * 30
                    a = 5000 * PI2 / r
                    .TX = ArrayOfTargets(AgentTarget(I)).X + Cos(a) * r
                    .TY = ArrayOfTargets(AgentTarget(I)).Y + Sin(a) * r

                Case 7            'Formation
                    dx = ArrayOfTargets(I).X
                    dy = ArrayOfTargets(I).Y

                    d = Sqr(dx * dx + dy * dy)
                    .cR = 0.65 + 0.35 * Cos(d * 1.25)
                    .cG = 0.65 + 0.35 * Cos(d * 0.5)
                    .cB = 0.65 + 0.35 * Sin(d * 1.1)

                End Select

                If MODE <> 7 Then .NReachedTargets = .NReachedTargets + 1&
            Else
                'Slow down when about to reach target
                dx = dx - .vX * VMUL * 1.33    ' 1.5
                dy = dy - .vY * VMUL * 1.33    ' 1.5
                d = dx * dx + dy * dy
                d = 1# / Sqr(d)
                dx = dx * d
                dy = dy * d

                .vX = .vX + dx * toTargetStrengthK * .maxV    '* .VTTS
                .vY = .vY + dy * toTargetStrengthK * .maxV    '* .VTTS


                '.VTTS = .VTTS * 0.999 + 0.001

                'control Max Speed
                d = .vX * .vX + .vY * .vY
                If d > .maxV2 Then
                    d = .maxV / Sqr(d)
                    .vX = .vX * d
                    .vY = .vY * d
                End If

            End If


            ''' [AAA]
            '...... Dont' Allow to move backwards
            Dim rX#, rY#
            '            If .VX * .ShoulderX + .VY * .ShoulderY < 0# Then  '0
            '            'If I = FOLLOW Then Stop
            '            D = Sqr(.VX * .VX + .VY * .VY)
            '            If .VX * .ShoulderY - .VY * .ShoulderX < 0 Then 'cross
            '            .VX = -.ShoulderY * D
            '            .VY = .ShoulderX * D
            '            Else
            '            .VX = .ShoulderY * D
            '            .VY = -.ShoulderX * D
            '            End If
            '            End If

            .X = .X + .vX
            .Y = .Y + .vY

            'keep in world
            If .X < 0# Then .X = 0#
            If .Y < 0# Then .Y = 0#
            If .X > WorldW Then .X = WorldW
            If .Y > WorldH Then .Y = WorldH

            'Preserve some V Change
            .VXchange = .VXchange * 0.5
            .VYchange = .VYchange * 0.5
            '            .VXchange = .VXchange * 0.33
            '            .VYchange = .VYchange * 0.33

            '            .V = GlobMaxSpeed2 * 0.5 + .VX * .VX + .VY * .VY

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
    Dim dx()      As Double
    Dim dy()      As Double
    Dim d()       As Double


    Dim CheckPosX As Double
    Dim CheckPosY As Double
    Dim dXX       As Double
    Dim dYY       As Double
    Dim NdX       As Double
    Dim NdY       As Double

    Dim DD        As Double
    Dim kSpeedI   As Double
    Dim kSpeedJ   As Double
    Dim AgentsDX  As Double
    Dim AgentsDY  As Double
    Dim JID       As Double

    Dim DOT_I#, DOT_J#

    Dim GlobAVOIDMUL As Double
    Dim KOtherSpeed As Double

    Dim Avoid     As Double
    Dim SqrDD     As Double
    Dim GoingToCollide As Double

    Dim ChangeMul1#, ChangeMul2#

    Ncollisions = 0

    GlobAVOIDMUL = GlobMaxSpeed * 0.125 * 0.85    '* 0.2    '* 5 * 2 '1.125
    KOtherSpeed = 0.1 * 2         '0.165           ' 0.22

    ' With Enabled "Power 2" and Smoothstep
    'GlobAVOIDMUL = GlobAVOIDMUL * 1.8

    ' With Enabled " Power 4 "
    GlobAVOIDMUL = GlobAVOIDMUL * 9 * 0.15 * 0.7    '' 0.16  ' 8 '7.2
    ' Without Smoothstep and " Power 4 " less collisioons

    With GRID
        .ResetPoints
        For I = 1 To NA
            .InsertPoint Agent(I).X, Agent(I).Y
        Next
        .GetPairsWDist P1(), P2(), dx(), dy(), d(), RVONpairs


        If CameraMode Then ReDim PairInfo(RVONpairs)

        For pair = 1 To RVONpairs

            I = P1(pair)
            J = P2(pair)

            CheckPosX = Agent(I).X + (Agent(I).vX - Agent(J).vX) * VMUL
            CheckPosY = Agent(I).Y + (Agent(I).vY - Agent(J).vY) * VMUL

            'CheckPosX = CheckPosX + (Agent(I).VXchange - Agent(J).VXchange) * VMUL
            'CheckPosY = CheckPosY + (Agent(I).VYchange - Agent(J).VYchange) * VMUL


            dXX = Agent(J).X - CheckPosX
            dYY = Agent(J).Y - CheckPosY

            DD = dXX * dXX + dYY * dYY

            '            DD = -sdOrientedVesica(vec3(dXX, dYY, 0), _
                         '                                   vec3(Agent(I).vX * VMUL + Agent(J).vX * VMUL, Agent(I).vY * VMUL + Agent(J).vY * VMUL, 0), _
                         '                                   vec3(-Agent(I).vX * VMUL - Agent(J).vX * VMUL, -Agent(I).vY * VMUL - Agent(J).vY * VMUL, 0), VMUL * 2)
            '
            '            If DD < 0 Then
            '                SqrDD = DD
            If DD < ControlDist2 Then
                SqrDD = Sqr(DD)

                'DD inverse proportional to DIST (Range 0,1)
                'dd = 1# - SqrDD * invControlDist
                GoingToCollide = 1# - SqrDD * invControlDist    '

                '                '                ' 'smoothstep
                '                '                GoingToCollide = GoingToCollide * GoingToCollide * (3# - 2# * GoingToCollide)
                '                ' "Power "
                ''                GoingToCollide = GoingToCollide * GoingToCollide
                ''                '                ' "Power 4"
                ''                GoingToCollide = GoingToCollide * GoingToCollide

                '                ' "Power 3"
                GoingToCollide = GoingToCollide * GoingToCollide * GoingToCollide

                'Normalize dXX,dYY and multiply by DD
                SqrDD = 1# / SqrDD
                dXX = dXX * SqrDD * GoingToCollide
                dYY = dYY * SqrDD * GoingToCollide

                ' More changes to Faster at this moment
                'kSpeedI = Agent(I).V / (Agent(I).V + Agent(J).V)
                ' More changes to Faster (MaxV) - Seems Better
                kSpeedI = Agent(I).maxV / (Agent(I).maxV + Agent(J).maxV)
                kSpeedJ = 1 - kSpeedI
                '

                'Compute Normalized DX DY of Agents I,J
                AgentsDX = dx(pair)
                AgentsDY = dy(pair)
                JID = 1# / Sqr(d(pair))
                AgentsDX = AgentsDX * JID
                AgentsDY = AgentsDY * JID

                '   Agent I   (DOT =  how much other Agent is in front (-1 to 1)
                DOT_I = Agent(I).ShoulderX * AgentsDX + Agent(I).ShoulderY * AgentsDY
                Avoid = Agent(I).Avoidance * (0.6 + DOT_I * 0.4)    'More Avoid if orher agent is in front
                '                ChangeMul1 = kSpeedI * toTargetStrengthK * GlobAvoidMUL * Avoid
                ChangeMul1 = kSpeedI * GlobAVOIDMUL * Avoid
                Agent(I).VXchange = Agent(I).VXchange - dXX * ChangeMul1
                Agent(I).VYchange = Agent(I).VYchange - dYY * ChangeMul1
                'It's useful to take some Velocity from other Agent
                If DOT_I > 0# Then
                    '                    DOT_I = DOT_I * GoingToCollide * GlobAVOIDMUL * KOtherSpeed
                    DOT_I = DOT_I ^ 0.5 * GoingToCollide * GlobAVOIDMUL * KOtherSpeed * kSpeedI
                    Agent(I).VXchange = Agent(I).VXchange + Agent(J).nVX * DOT_I    '- Agent(I).VX * DOT_J
                    Agent(I).VYchange = Agent(I).VYchange + Agent(J).nVY * DOT_I    '- Agent(I).VY * DOT_J
                Else
                    DOT_I = 0
                End If

                '   Agent J   (DOT =  how much other Agent is in front (-1 to 1)
                DOT_J = Agent(J).ShoulderX * -AgentsDX + Agent(J).ShoulderY * -AgentsDY
                Avoid = Agent(J).Avoidance * (0.6 + DOT_J * 0.4)
                '                ChangeMul2 = kSpeedJ * toTargetStrengthK * GlobAvoidMUL * Avoid
                ChangeMul2 = kSpeedJ * GlobAVOIDMUL * Avoid
                Agent(J).VXchange = Agent(J).VXchange + dXX * ChangeMul2
                Agent(J).VYchange = Agent(J).VYchange + dYY * ChangeMul2
                'It's useful to take some Velocity from other Agent
                If DOT_J > 0# Then
                    '                    DOT_J = DOT_J * GoingToCollide * GlobAVOIDMUL * KOtherSpeed
                    DOT_J = DOT_J ^ 0.5 * GoingToCollide * GlobAVOIDMUL * KOtherSpeed * kSpeedJ
                    Agent(J).VXchange = Agent(J).VXchange + Agent(I).nVX * DOT_J    '- Agent(J).VX * DOT_J
                    Agent(J).VYchange = Agent(J).VYchange + Agent(I).nVY * DOT_J    '- Agent(J).VY * DOT_J
                Else
                    DOT_J = 0
                End If
                '---------------------------------------------------------------




                If CameraMode Then
                    With PairInfo(pair)
                        .Enabled = True
                        .a = I: .b = J
                        .aVX = Agent(I).vX * VMUL
                        .aVY = Agent(I).vY * VMUL
                        .bVX = Agent(J).vX * VMUL
                        .bVY = Agent(J).vY * VMUL
                        .chkX = CheckPosX
                        .chkY = CheckPosY
                        .chkX2 = Agent(J).X + (Agent(J).vX - Agent(I).vX) * VMUL
                        .chkY2 = Agent(J).Y + (Agent(J).vY - Agent(I).vY) * VMUL
                        .GTC = GoingToCollide
                        .aVXch = -dXX * ChangeMul1 + Agent(J).nVX * DOT_I
                        .aVYch = -dYY * ChangeMul1 + Agent(J).nVY * DOT_I
                        .bVXch = dXX * ChangeMul2 + Agent(I).nVX * DOT_J
                        .bVYch = dYY * ChangeMul2 + Agent(I).nVY * DOT_J
                        ''                        .aVXch = -dXX * ChangeMul1 + dVX * DOT_I
                        ''                        .aVYch = -dYY * ChangeMul1 + dVY * DOT_I
                        ''                        .bVXch = dXX * ChangeMul2 - dVX * DOT_J
                        ''                        .bVYch = dYY * ChangeMul2 - dVY * DOT_J
                    End With

                End If


                ''ChangeMul1 = Max(0.1, 1 - ChangeMul1 * 0.0005)
                ''ChangeMul2 = Max(0.1, 1 - ChangeMul2 * 0.0005)
                ''Agent(I).VTTS = Agent(I).VTTS * ChangeMul1
                ''Agent(J).VTTS = Agent(J).VTTS * ChangeMul2


            End If


            '------------------------------------
            'Separate


            If d(pair) < AgentsMinDist2 Then
                '          Stop
                '            FOLLOW = I

                DD = Sqr(d(pair))
                dXX = dx(pair) / DD
                dYY = dy(pair) / DD
                dXX = dXX * (AgentsMinDist - DD) * 0.5
                dYY = dYY * (AgentsMinDist - DD) * 0.5
                Agent(I).X = Agent(I).X - dXX
                Agent(I).Y = Agent(I).Y - dYY
                Agent(J).X = Agent(J).X + dXX
                Agent(J).Y = Agent(J).Y + dYY
                Ncollisions = Ncollisions + 1
            End If


        Next
    End With

    If (CNT And 7&) = 0 Then fMain.Caption = "Click to change Mode.   Mode: " & MODE & "         Follow: " & FollowDesc & "     Right click to change Camera follow  " & "   Pairs: " & RVONpairs & "   Collisions: " & Ncollisions & "              " & Format((Frame / 25) / (86000), "HH:MM:SS")

End Sub

Public Sub MAINLOOP()
    doLOOP = True


    Do

        MOVE

        'If (CNT And 1&) = 0& Then DRAW
        If (CNT And 1&) = 0& Then

            If Not (CameraMode) Then draw_CAMERA Else: Draw_PAIR
        End If


        If (CNT And 31&) = 0& Then
            If FollowMode = 9 Then FOLLOW = FOLLOWHiTarg: FollowDesc = "MostTargetReached " & FOLLOW
        End If


        '        If (CNT And 3&) = 0 Then
        '            updateTRAILS
        '        End If


        If doFRAMES Then
            If (CNT And 1&) = 0 Then
                SRF.WriteContentToPngFile App.Path & "\FRAMES\" & Format(Frame, "00000") & ".png"
                Frame = Frame + 1
            End If
        End If



        CNT = CNT + 1
        DoEvents

    Loop While doLOOP

End Sub

Private Sub updateTRAILS()
    Dim I         As Long
    For I = 1 To NA
        Agent(I).SciaX(TRAILS) = Agent(I).X
        Agent(I).SciaY(TRAILS) = Agent(I).Y
    Next

    TRAILS = (TRAILS + 1) Mod (Ntrail + 1)

End Sub

Public Function FOLLOWworstRule() As Long
    Dim r         As Double
    Dim I         As Long
    r = 99
    For I = 1 To NA
        If Agent(I).Avoidance < r Then
            r = Agent(I).Avoidance: FOLLOWworstRule = I
        End If
    Next
End Function
Public Function FOLLOWSlower() As Long
    Dim r         As Double
    Dim I         As Long
    r = 99
    For I = 1 To NA
        If Agent(I).maxV < r Then
            r = Agent(I).maxV: FOLLOWSlower = I
        End If
    Next
End Function
Public Function FOLLOWworstFaster() As Long
    Dim r         As Double
    Dim I         As Long
    r = -999
    For I = 1 To NA
        If Agent(I).maxV - Agent(I).Avoidance > r Then
            r = Agent(I).maxV - Agent(I).Avoidance: FOLLOWworstFaster = I
        End If
    Next
End Function
Public Function FOLLOWworstSlower() As Long
    Dim r         As Double
    Dim I         As Long
    r = -999
    For I = 1 To NA
        If -Agent(I).maxV - Agent(I).Avoidance > r Then
            r = -Agent(I).maxV - Agent(I).Avoidance: FOLLOWworstSlower = I
        End If
    Next
End Function
Public Function FOLLOWbestRule() As Long
    Dim r         As Double
    Dim I         As Long
    r = 0
    For I = 1 To NA
        If Agent(I).Avoidance > r Then
            r = Agent(I).Avoidance: FOLLOWbestRule = I
        End If
    Next
End Function
Public Function FOLLOWFaster() As Long
    Dim r         As Double
    Dim I         As Long
    r = 0
    For I = 1 To NA
        If Agent(I).maxV > r Then
            r = Agent(I).maxV: FOLLOWFaster = I
        End If
    Next
End Function
Public Function FOLLOWbestFaster() As Long
    Dim r         As Double
    Dim I         As Long
    r = 0
    For I = 1 To NA
        If Agent(I).Avoidance + Agent(I).maxV > r Then
            r = Agent(I).Avoidance + Agent(I).maxV: FOLLOWbestFaster = I
        End If
    Next
End Function

Public Function FOLLOWbestSlower() As Long
    Dim r         As Double
    Dim I         As Long
    r = -99999
    For I = 1 To NA
        If Agent(I).Avoidance - Agent(I).maxV > r Then
            r = Agent(I).Avoidance - Agent(I).maxV: FOLLOWbestSlower = I
        End If
    Next
End Function
Public Function FOLLOWHiTarg() As Long
    Dim I         As Long
    For I = 1 To NA
        If Agent(I).NReachedTargets > MAXTR Then
            MAXTR = Agent(I).NReachedTargets: FOLLOWHiTarg = I
        End If
    Next
    If FOLLOWHiTarg = 0 Then FOLLOWHiTarg = FOLLOW
End Function



Private Sub BUILDHuman(I As Long)
    Dim L         As tCapsule

    Dim X#, Y#
    Dim nX#, nY#
    Dim dx#, dy#: Dim a#, A2#
    Dim w#: Dim h#


    Dim Ank       As tVec3
    Dim Feet      As tVec3
    Dim Knee      As tVec3
    Dim tmp1      As tVec3
    Dim tmp2      As tVec3

    X = Agent(I).X
    Y = Agent(I).Y
    dx = Agent(I).ShoulderX
    dy = Agent(I).ShoulderY
    w = Agent(I).Walked

    nX = Agent(I).nVX
    nY = Agent(I).nVY
    With L                        'BODY
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
        'TransformLine .P1, .P2, DX, DY, X, Y:        ADDlineToScreen L, I
        .Size = 1
        '------------------------------------------------------------------------

        'LEGS
        ''        A = W * 0.58    '0.5
        ''        .P1 = vec3(0, -4.5, 1.25)
        ''        H = Sin(A) * 2.5: If H > 0 Then H = 0
        ''        .P2 = vec3(Cos(A) * 3, H, 1.25)
        ''        TransformLine .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I
        ''
        ''        A = A + PI
        ''        .P1 = vec3(0, -4.5, -1.25)
        ''        H = Sin(A) * 2.5: If H > 0 Then H = 0
        ''        .P2 = vec3(Cos(A) * 3, H, -1.25)
        ''        TransformLine .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I
        ''



        a = w * 0.65              '0.5
        .P1 = vec3(0, -4.5, 1.25)
        h = Sin(a) * 1.5: If h > 0 Then h = 0
        .P2 = vec3(Cos(a) * 3, h, 1.25)


        '        tmp2.X = FEETs(I).Lfeet.X - X
        '        tmp2.Z = FEETs(I).Lfeet.Y - X
        '        .P2.X = tmp2.X * nX + tmp2.Y * nX
        '        .P2.Z = tmp2.Z * nY - tmp2.Z * nY
        '        .P2.Y = 0
        '


        IKSolve .P1, .P2, 2.5, 2, Knee, tmp1
        If dx * nX + dy * nY <= 0# Then Knee = tmp1
        tmp1 = .P1
        tmp2 = .P2
        Knee.Z = .P1.Z

        .P2 = Knee
        TransformLine .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I
        .P1 = Knee
        .P2 = tmp2
        TransformLine .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I


        a = a + PI
        .P1 = vec3(0, -4.5, -1.25)
        h = Sin(a) * 1.6: If h > 0 Then h = 0
        .P2 = vec3(Cos(a) * 3, h, -1.25)



        IKSolve .P1, .P2, 2.5, 2, Knee, tmp1
        tmp1 = .P1
        tmp2 = .P2
        Knee.Z = .P1.Z
        .P2 = Knee
        TransformLine .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I
        .P1 = Knee
        .P2 = tmp2
        TransformLine .P1, .P2, nX, nY, X, Y: ADDlineToScreen L, I
        '------------------------------------------------------------------------

        'ARMS
        .Size = 0.8
        '        .P1 = vec3(0, -8, 1.5)
        '        .P2 = vec3(Cos(A) + 0.5, -4 - Cos(A), 3)
        '        TransformLine .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I
        '
        '        .P1 = vec3(0, -8, -1.5)
        '        .P2 = vec3(Cos(A + PI) + 0.5, -4 - Cos(A + PI), -3)
        '        TransformLine .P1, .P2, DX, DY, X, Y: ADDlineToScreen L, I

        A2 = Cos(a) * 1.4
        .P1 = vec3(0, -7.5, 1.65)
        .P2 = vec3(A2 + 0.5, -3.6 - Abs(A2) * 0.77, 2.65)
        TransformLine .P1, .P2, dx, dy, X, Y: ADDlineToScreen L, I

        .P1 = vec3(0, -7.5, -1.65)
        .P2 = vec3(-A2 + 0.5, -3.6 - Abs(A2) * 0.77, -2.65)
        TransformLine .P1, .P2, dx, dy, X, Y: ADDlineToScreen L, I


        'HEAD
        '        .P1 = vec3(0.2, -9.5, 0)
        '        .P2 = vec3(0.4, -10, 0)
        .P1 = vec3(0.45, -9.5 + 0.5, 0)
        .P2 = vec3(0.8, -9.5 + 0.5, 0)
        ' TransformLine .P1, .P2, nX, nY, X, Y: L.Size = 1.75: ADDlineToScreen L, I
        TransformLine .P1, .P2, dx, dy, X, Y: L.Size = 1.75: ADDlineToScreen L, I





        'TARGET
        .P1.X = Agent(I).TX
        .P1.Z = Agent(I).TY
        .P1.Y = 0
        .P2 = L.P1
        .P2.Y = -0.2              '-0.1
        .Size = 1.33              '1.5
        ADDlineToScreen L, I
    End With
End Sub

'Private Sub RotateXZ(p As tVec3, cosA#, sinA#)
'    Dim c         As tVec3: c = p
'    p.X = c.X * cosA + c.Z * -sinA
'    p.Z = c.X * sinA + c.Z * cosA
'End Sub
'Private Sub AddXZ(p As tVec3, X#, Y#)
'    p.X = p.X + X
'    p.Z = p.Z + Y
'End Sub

Private Sub TransformLine(P1 As tVec3, P2 As tVec3, CosA#, SinA#, X#, Y#)
    '    RotateXZ P1, cosA, sinA
    '    RotateXZ P2, cosA, sinA
    '    AddXZ P1, X, Y
    '    AddXZ P2, X, Y
    Dim c         As tVec3: c = P1
    P1.X = c.X * CosA + c.Z * -SinA + X
    P1.Z = c.X * SinA + c.Z * CosA + Y

    c = P2
    P2.X = c.X * CosA + c.Z * -SinA + X
    P2.Z = c.X * SinA + c.Z * CosA + Y

End Sub


Public Sub INIT_Targets()
    Dim I         As Long
    Dim FirstTarg As Boolean
    Dim r         As Double
    Dim a         As Double

    If NTargets = 0 Then
        ReDim AgentTarget(NA)
        FirstTarg = True

    End If

    NTargets = 3 + 2
    ReDim ArrayOfTargets(NTargets)


    For I = 0 To NTargets
        ArrayOfTargets(I).X = 100 + Rnd * (WorldW - 200)
        ArrayOfTargets(I).Y = 100 + Rnd * (WorldH - 200)
    Next

    If FirstTarg Then
        '        For I = 1 To NA
        '            R = 53
        '            A = I / NA * PI
        '            Agent(I).TX = ArrayOfTargets(0).X + Cos(A) * R
        '            Agent(I).TY = ArrayOfTargets(0).Y + Sin(A) * R
        '        Next

        For I = 1 To NA
            '            R = 20 + 99 * I / NA
            r = I / NA: r = Sqr(r)
            '            R = 10 + 99 * R
            r = 4 + r * 30

            a = 5000 * PI2 / r
            Agent(I).TX = ArrayOfTargets(0).X + Cos(a) * r
            Agent(I).TY = ArrayOfTargets(0).Y + Sin(a) * r
        Next
    End If

End Sub


























'Oriented Vesica - exact   (https://www.shadertoy.com/view/cs2yzG)
'
'float sdOrientedVesica( vec2 p, vec2 a, vec2 b, float w )
'{
'    float r = 0.5*length(b-a);
'    float d = 0.5*(r*r-w*w)/w;
'    vec2 v = (b-a)/r;
'    vec2 c = (b+a)*0.5;
'    vec2 q = 0.5*abs(mat2(v.y,v.x,-v.x,v.y)*(p-c));
'    vec3 h = (r*q.x<d*(q.y-r)) ? vec3(0.0,r,0.0) : vec3(-d,0.0,d+w);
'    return length( q-h.xy) - h.z;
'}
''Oriented Vesica - exact   (https://www.shadertoy.com/view/cs2yzG)

Private Function sdOrientedVesica(p As tVec3, a As tVec3, b As tVec3, w As Double) As Double
    Dim r#, d#
    Dim vX#, vY#
    Dim cX#, cY#
    Dim qX#, qY#
    Dim h         As tVec3
    Dim dx        As Double
    Dim dy        As Double

    dx = b.X - a.X
    dy = b.Y - a.Y

    r = 0.5 * Sqr(dx * dx + dy * dy)
    d = 0.5 * (r * r - w * w) / w
    vX = dx / r
    vY = dy / r
    cX = (b.X + a.X) * 0.5
    cY = (b.Y + a.Y) * 0.5

    dx = p.X - cX
    dy = p.X - cX

    qX = 0.5 * Abs(dx * vY + dy * vX)
    qY = 0.5 * Abs(dx * -vX + dy * vY)

    '    vec2 q = 0.5*abs(mat2(v.y,v.x,-v.x,v.y)*(p-c));
    '    vec3 h = (r*q.x<d*(q.y-r)) ? vec3(0.0,r,0.0) : vec3(-d,0.0,d+w);
    '    return length( q-h.xy) - h.z;
    If (r * qX < d * (qY - r)) Then
        h = vec3(0#, r, 0#)
    Else
        h = vec3(-d, 0#, d + w)
    End If

    sdOrientedVesica = Sqr(dx * dx + dy * dy) - h.Z

End Function


Private Sub OUTINFO()
    'https://www.vbforums.com/showthread.php?898105-RESOLVED-Fugly-Font-itis-Is-there-a-better-way-to-Cairo-this&p=5583735&viewfull=1#post5583735
    CC.SetLineWidth 4             'Stroke width
    CC.SelectFont "Segoe UI", 12, RGB(255, 255, 255)    'Font selection
    CC.TextOut 20, 15, FollowDesc & "      Mode: " & MODE, , , True    'Path for stroke

    '    CC.SetSourceColor vbBlack
    '    CC.Stroke True    ' Keep the path open so we can perform a subsequent fill
    CC.SetSourceColor vbWhite
    CC.Fill                       ' Fill Close the path

End Sub
