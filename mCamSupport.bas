Attribute VB_Name = "mCamSupport"
' Minimal vec3 math for c3DEasyCam.cls
Option Explicit

Public Const PI   As Double = 3.14159265358979
Public Const PIh  As Double = 1.5707963267949
Public Const PI2  As Double = 6.28318530717959
Public Const InvPI As Double = 1 / 3.14159265358979
Public Const InvPIh As Double = 1 / 1.5707963267949
Public Const InvPI2 As Double = 1 / 6.28318530717959

Public Type tVec3
    X             As Double
    Y             As Double
    Z             As Double
End Type

Public Type tCapsule
    P1            As tVec3
    P2            As tVec3
    ScreenP1      As tVec3
    ScreenP2      As tVec3
    R             As Double
    AgentIndex    As Long
    Size          As Double
    CamD          As Double

End Type

Public CAPSULES() As tCapsule
Public NCapsulesInScreen As Long
Public MaxNCapsulesInScreen As Long

Public CAM        As c3DEasyCam



Public Function vec3(X As Double, Y As Double, Z As Double) As tVec3
    With vec3
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Function Normalize3(V As tVec3) As tVec3
    Dim D         As Double

    With V
        D = (.X * .X + .Y * .Y + .Z * .Z)
        If D Then
            D = 1# / Sqr(D)
            Normalize3.X = .X * D
            Normalize3.Y = .Y * D
            Normalize3.Z = .Z * D
        Else
            '        MsgBox "??? NORMALIZE"
        End If
    End With

End Function
Public Function DIFF3(V1 As tVec3, V2 As tVec3) As tVec3
    With DIFF3
        .X = V1.X - V2.X
        .Y = V1.Y - V2.Y
        .Z = V1.Z - V2.Z
    End With
End Function

Public Function CROSS3(A As tVec3, B As tVec3) As tVec3
    With A
        CROSS3.X = .Y * B.Z - .Z * B.Y
        CROSS3.Y = .Z * B.X - .X * B.Z
        CROSS3.Z = .X * B.Y - .Y * B.X
    End With
End Function
Public Function DOT3(V1 As tVec3, V2 As tVec3) As Double

    DOT3 = (V1.X * V2.X) + _
           (V1.Y * V2.Y) + _
           (V1.Z * V2.Z)

End Function
Public Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
    If X Then
        '        Atan2 = -PI + Atn(Y / X) - (X > 0#) * Pi
        Atan2 = Atn(Y / X) + PI * (X < 0#)
    Else
        Atan2 = -PIh - (Y > 0#) * PI
    End If
    ' While Atan2 < 0: Atan2 = Atan2 + PI2: Wend
    ' While Atan2 > PI2: Atan2 = Atan2 - PI2: Wend
End Function
Public Function Length3(V As tVec3) As Double
    With V
        Length3 = Sqr(.X * .X + .Y * .Y + .Z * .Z)
    End With
End Function

Public Function Length32(V As tVec3) As Double
    With V
        Length32 = .X * .X + .Y * .Y + .Z * .Z
    End With
End Function

Public Function MUL3(V As tVec3, A As Double) As tVec3
    With V
        MUL3.X = .X * A
        MUL3.Y = .Y * A
        MUL3.Z = .Z * A
    End With
End Function

Public Function MUL3V(V As tVec3, V2 As tVec3) As tVec3
    With V
        MUL3V.X = .X * V2.X
        MUL3V.Y = .Y * V2.Y
        MUL3V.Z = .Z * V2.Z
    End With
End Function

Public Function SUM3(V1 As tVec3, V2 As tVec3) As tVec3
    With V1
        SUM3.X = .X + V2.X
        SUM3.Y = .Y + V2.Y
        SUM3.Z = .Z + V2.Z
    End With
End Function

Public Function SUM3v(V As tVec3, A As Double) As tVec3
    With V
        SUM3v.X = .X + A
        SUM3v.Y = .Y + A
        SUM3v.Z = .Z + A
    End With
End Function

Public Function RayPlaneIntersect(rayVector As tVec3, rayPoint As tVec3, PlaneNormal As tVec3, planePoint As tVec3) As tVec3
    'https://rosettacode.org/wiki/Find_the_intersection_of_a_line_with_a_plane#C.23

    Dim Diff      As tVec3
    Dim prod1     As Double
    Dim prod2     As Double
    Dim prod3     As Double

    Diff = DIFF3(rayPoint, planePoint)
    prod1 = DOT3(Diff, PlaneNormal)
    prod2 = DOT3(rayVector, PlaneNormal)
    prod3 = prod1 / prod2
    RayPlaneIntersect = DIFF3(rayPoint, MUL3(rayVector, prod3))

End Function


'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------


' OTHER FUNCTION - NOT used by CAMERA

Public Function Max(A As Double, B As Double) As Double
    Max = A
    If B > Max Then Max = B
End Function
Public Function min(A As Double, B As Double) As Double
    min = A
    If B < min Then min = B
End Function

''https://iquilezles.org/articles/simpleik/
'Public Function IKSolve(ByRef C1 As tVec3, _
 '                        ByRef C2 As tVec3, _
 '                        ByRef R1 As Double, _
 '                        ByRef R2 As Double, _
 '                        ByRef Sol1 As tVec3, _
 '                        ByRef Sol2 As tVec3) As Boolean
''
''{
''    float h = dot(p,p);
''    float w = h + r1*r1 - r2*r2;
''    float s = max(4.0*r1*r1*h - w*w,0.0);
''    return (w*p + vec2(-p.y,p.x)*sqrt(s)) * 0.5/h;
''}
'    Dim p         As tVec3
'    Dim H#, W#, S#
'
'    p = DIFF3(C1, C2)
'    p.Z = 0
'
'    H = DOT3(p, p)
'    W = H + R1 * R1 - R2 * R2
'    S = 4# * R1 * R1 * H - W * W: If S < 0# Then S = 0#
'
'    S = Sqr(S)
'    H = 0.5 / H
'
'    Sol1.X = (p.X * W - p.Y * S) * H
'    Sol1.Y = (p.Y * W + p.X * S) * H
'    If Sol1.X > -0.99 Then IKSolve = True
'    Sol1.X = Sol1.X + C2.X
'    Sol1.Y = Sol1.Y + C2.Y
'
'
'    Sol2.X = (p.X * W + p.Y * S) * H
'    Sol2.Y = (p.Y * W - p.X * S) * H
'    If Sol2.X > -0.99 Then IKSolve = True
'    Sol2.X = Sol2.X + C2.X
'    Sol2.Y = Sol2.Y + C2.Y
'
'    IKSolve = True
'
'
'End Function

'https://iquilezles.org/articles/simpleik/
Public Function IKSolve(ByRef Origin As tVec3, _
                        ByRef Target As tVec3, _
                        ByRef R1 As Double, _
                        ByRef R2 As Double, _
                        ByRef Sol1 As tVec3, _
                        ByRef Sol2 As tVec3) As Boolean
    ' NOTE: Target will change position if no solution found!
    Dim pX#, pY#
    Dim H#, W#, S#

    pX = Origin.X - Target.X
    pY = Origin.Y - Target.Y

    H = pX * pX + pY * pY
    W = H + R2 * R2 - R1 * R1
    S = 4# * R2 * R2 * H - W * W

    If S > 0# Then
        S = Sqr(S)
        H = 0.5 / H
        Sol1.X = Target.X + (pX * W - pY * S) * H
        Sol1.Y = Target.Y + (pY * W + pX * S) * H
        Sol2.X = Target.X + (pX * W + pY * S) * H
        Sol2.Y = Target.Y + (pY * W - pX * S) * H
        IKSolve = True
    Else
        'No solution (so Move target and get Solutions [by reexre])
        H = 1# / Sqr(H)
        pX = pX * H: pY = pY * H
        Sol1.X = Origin.X - pX * R1
        Sol1.Y = Origin.Y - pY * R1
        Sol2 = Sol1
        If W > 0# Then            'Target too far
            Target.X = Origin.X - pX * (R1 + R2)
            Target.Y = Origin.Y - pY * (R1 + R2)
        Else                      'Target too close to Origin
            Target.X = Sol1.X + pX * R2
            Target.Y = Sol1.Y + pY * R2
        End If
    End If
    '{
    '    float h = dot(p,p);
    '    float w = h + r2*r2 - r1*r1;
    '    float s = max(4.0*r2*r2*h - w*w,0.0);
    '    return (w*p + vec2(-py,px)*sqrt(s)) * 0.5/h;
    '}
End Function



'SEGMETS**********************************
Public Sub IntersectOfLines(LA1x#, LA1y#, LA2x#, LA2y#, _
                            LB1x#, LB1y#, LB2x#, LB2y#, rX#, rY#)

    Dim D         As Double
    Dim NA As Double: Dim NB As Double
    Dim Dx1 As Double: Dim Dx2 As Double
    Dim DY1 As Double: Dim Dy2 As Double
    Dim uA As Double: Dim uB As Double

    Dx1 = LA2x - LA1x
    DY1 = LA2y - LA1y
    Dx2 = LB2x - LB1x
    Dy2 = LB2y - LB1y

    ' Denominator for ua and ub are the sP1me, so store this calculation
    D = (Dy2) * (Dx1) - _
        (Dx2) * (DY1)


    ' Make sure there is not a division by zero - this also indicates that
    ' the lines are parallel.
    ' If NA and NB were both equal to zero the lines would be on top of each
    ' other (coincidental).  This check is not done because it is not
    ' necessP1ry for this implementation (the parallel check accounts for this).
    rX = -9999
    rY = -9999

    If D = 0 Then Exit Sub
    D = 1# / D

    'NA and NB are calculated as seperate values for readability
    NA = (Dx2) * (LA1y - LB1y) - _
         (Dy2) * (LA1x - LB1x)



    ' Calculate the intermediate fractional point that the lines potentially intersect.
    uA = NA * D

    ' The fractional point will be between 0 and 1 inclusive if the lines
    ' intersect.  If the fractional calculation is larger than 1 or smaller
    ' than 0 the lines would need to be longer to intersect.
    If uA >= 0# Then
        If uA <= 1# Then
            ' Calculate the intermediate fractional point that the lines potentially intersect.
            NB = (Dx1) * (LA1y - LB1y) - _
                 (DY1) * (LA1x - LB1x)
            uB = NB * D
            If uB >= 0# Then
                If uB <= 1# Then
                    rX = LA1x + (uA * Dx1)
                    rY = LA1y + (uA * DY1)
                End If
            End If
        End If
    End If
End Sub






Private Function BnearestThanA(A As tCapsule, B As tCapsule) As Boolean

    Dim D         As Double
    Dim NA As Double: Dim NB As Double
    Dim Dx1 As Double: Dim Dx2 As Double
    Dim DY1 As Double: Dim Dy2 As Double
    Dim uA As Double: Dim uB As Double


    '    BnearestThanA = A.ScreenP1.Z + A.ScreenP2.Z < B.ScreenP1.Z + B.ScreenP2.Z
    '    BnearestThanA = min(A.ScreenP1.Z, A.ScreenP2.Z) < min(B.ScreenP1.Z, B.ScreenP2.Z)
    BnearestThanA = A.CamD > B.CamD


    Exit Function

    Const E       As Double = 0.001
    If Abs(A.ScreenP1.X - B.ScreenP1.X) < E And Abs(A.ScreenP1.Y - B.ScreenP1.Y) < E Then
        BnearestThanA = A.ScreenP2.Z > B.ScreenP2.Z: Exit Function
    ElseIf Abs(A.ScreenP1.X - B.ScreenP2.X) < E And Abs(A.ScreenP1.Y - B.ScreenP2.Y) < E Then
        BnearestThanA = A.ScreenP2.Z > B.ScreenP1.Z: Exit Function
    ElseIf Abs(A.ScreenP2.X - B.ScreenP1.X) < E And Abs(A.ScreenP2.Y - B.ScreenP1.Y) < E Then
        BnearestThanA = A.ScreenP1.Z > B.ScreenP2.Z: Exit Function
    ElseIf Abs(A.ScreenP2.X - B.ScreenP2.X) < E And Abs(A.ScreenP2.Y - B.ScreenP2.Y) < E Then
        BnearestThanA = A.ScreenP1.Z > B.ScreenP1.Z: Exit Function

    Else

        Dx1 = A.ScreenP2.X - A.ScreenP1.X
        DY1 = A.ScreenP2.Y - A.ScreenP1.Y
        Dx2 = B.ScreenP2.X - B.ScreenP1.X
        Dy2 = B.ScreenP2.Y - B.ScreenP1.Y

        ' Denominator for ua and ub are the sP1me, so store this calculation
        D = (Dy2) * (Dx1) - _
            (Dx2) * (DY1)


        '        ' Make sure there is not a division by zero - this also indicates that
        '        ' the lines are parallel.
        '        ' If NA and NB were both equal to zero the lines would be on top of each
        '        ' other (coincidental).  This check is not done because it is not
        '        ' necessP1ry for this implementation (the parallel check accounts for this).
        '        IntersectOfLinesV3.X = -9999
        '        IntersectOfLinesV3.Y = -9999

        If D = 0 Then Exit Function
        D = 1# / D

        'NA and NB are calculated as seperate values for readability
        NA = (Dx2) * (A.ScreenP1.Y - B.ScreenP1.Y) - _
             (Dy2) * (A.ScreenP1.X - B.ScreenP1.X)


        ' Calculate the intermediate fractional point that the lines potentially intersect.
        uA = NA * D

        ' The fractional point will be between 0 and 1 inclusive if the lines
        ' intersect.  If the fractional calculation is larger than 1 or smaller
        ' than 0 the lines would need to be longer to intersect.
        If uA >= 0# Then
            If uA <= 1# Then
                ' Calculate the intermediate fractional point that the lines potentially intersect.
                NB = (Dx1) * (A.ScreenP1.Y - B.ScreenP1.Y) - _
                     (DY1) * (A.ScreenP1.X - B.ScreenP1.X)

                uB = NB * D
                If uB >= 0# Then
                    If uB <= 1# Then
                        Stop
                        '                        IntersectOfLinesV3.X = A.ScreenP1.X + (uA * Dx1)
                        '                        IntersectOfLinesV3.Y = A.ScreenP1.Y + (uA * DY1)
                        ' Z ?
                        BnearestThanA = A.ScreenP1.Z + (uA * (A.ScreenP2.Z - A.ScreenP1.Z)) > _
                                        B.ScreenP1.Z + (uB * (B.ScreenP2.Z - B.ScreenP1.Z))    'Z of line B
                        Exit Function
                    End If
                End If
            End If
        End If
    End If


    'If DefTrue Then BnearestThanA = True

End Function

Public Sub ADDlineToScreen(L As tCapsule, I As Long)
    Dim Vis       As Boolean
    CAM.FarPlane = 5000
    With L
        CAM.LineToScreen .P1, .P2, .ScreenP1, .ScreenP2, Vis
        'Distance from camera (z) of screenP is inverse 1/
        If Vis Then
            If .Size = 0 Then .Size = 1
            NCapsulesInScreen = NCapsulesInScreen + 1
            If NCapsulesInScreen > MaxNCapsulesInScreen Then
                MaxNCapsulesInScreen = NCapsulesInScreen * 2
                ReDim Preserve CAPSULES(MaxNCapsulesInScreen)
            End If
            .AgentIndex = I
            CAPSULES(NCapsulesInScreen) = L
            '            CAPSULES(NCapsulesInScreen).CamD = 1# / (0.5 * (.ScreenP1.Z + .ScreenP2.Z))
            CAPSULES(NCapsulesInScreen).CamD = 1# / (Max(.ScreenP1.Z, .ScreenP2.Z)) - .Size * 0.5

        End If


        If RunningNotInIDE Then   'Compiled (not IDE)
            '    '''---------- SHADOWS
            CAM.FarPlane = 400
            Vis = False
            CAM.LineToScreen vec3(.P1.X + .P1.Y * 1.3, 0, .P1.Z + .P1.Y * 1.2), _
                             vec3(.P2.X + .P2.Y * 1.3, 0, .P2.Z + .P2.Y * 1.2), .ScreenP1, .ScreenP2, Vis

            If Vis Then
                If L.Size = 0 Then .Size = 1
                NCapsulesInScreen = NCapsulesInScreen + 1
                If NCapsulesInScreen > MaxNCapsulesInScreen Then
                    MaxNCapsulesInScreen = NCapsulesInScreen * 2
                    ReDim Preserve CAPSULES(MaxNCapsulesInScreen)
                End If
                .AgentIndex = -1
                .ScreenP1.Z = .ScreenP1.Z * 0.5
                .ScreenP2.Z = .ScreenP2.Z * 0.5

                CAPSULES(NCapsulesInScreen) = L
                '                CAPSULES(NCapsulesInScreen).CamD = 2 * (1# / (0.5 * (.ScreenP1.Z + .ScreenP2.Z)))
                CAPSULES(NCapsulesInScreen).CamD = 2 * (1# / (Max(.ScreenP1.Z, .ScreenP2.Z)) - .Size * 0.5)
            End If
            '    --------------------------
        End If
    End With


End Sub
'Private Sub QuickSortFaces(List() As tFace, ByVal min As Long, ByVal Max As Long)
'
'' FROM HI to LOW  'https://www.vbforums.com/showthread.php?11192-quicksort
'    Dim Low As Long, high As Long, temp As tFace, TestElement As Double
'    'Debug.Print min, max
'    Low = min: high = Max
'    TestElement = List((min + Max) / 2).DistToCam
'    Do
'        Do While List(Low).DistToCam > TestElement: Low = Low + 1&: Loop
'        Do While List(high).DistToCam < TestElement: high = high - 1&: Loop
'        If (Low <= high) Then
'            temp = List(Low): List(Low) = List(high): List(high) = temp
'            Low = Low + 1&: high = high - 1&
'            '            Swaps = Swaps + 1
'        End If
'    Loop While (Low <= high)
'    If (min < high) Then QuickSortFaces List, min, high
'    If (Low < Max) Then QuickSortFaces List, Low, Max
'End Sub

Public Sub QuickSortCapsules(List() As tCapsule, ByVal min As Long, ByVal Max As Long)

    ' FROM HI to LOW  'https://www.vbforums.com/showthread.php?11192-quicksort
    Dim Low As Long, high As Long, temp As tCapsule, TestCapsule As tCapsule
    Dim TestDist#
    'Debug.Print min, max
    Low = min: high = Max
    'TestCapsule = List((min + Max) / 2)
    TestDist = (List(min).CamD + List(Max).CamD) * 0.5
    Do
        '        '        Do While List(Low).DistToCam > TestCapsule: Low = Low + 1&: Loop
        '        '        Do While List(high).DistToCam < TestCapsule: high = high - 1&: Loop
        '
        '        Do While Low < Max And (BnearestThanA(List(Low), TestCapsule)): Low = Low + 1&: Loop
        '        Do While high > min And Not (BnearestThanA(List(high), TestCapsule)): high = high - 1&: Loop

        Do While (List(Low).CamD > TestDist): Low = Low + 1&: Loop
        Do While (List(high).CamD < TestDist): high = high - 1&: Loop


        If (Low <= high) Then
            temp = List(Low): List(Low) = List(high): List(high) = temp
            Low = Low + 1&: high = high - 1&
            '            Swaps = Swaps + 1
        End If
    Loop While (Low <= high)
    If (min < high) Then QuickSortCapsules List, min, high
    If (Low < Max) Then QuickSortCapsules List, Low, Max


End Sub



Public Function VectorReflect(VX#, VY#, WallX#, WallY#, rX#, rY#)
    'Function returning the reflection of one vector around another.
    'it's used to calculate the rebound of a Vector on another Vector
    'Vector "V" represents current velocity of a point.
    'Vector "Wall" represent the angle of a wall where the point Bounces.
    'Returns the vector velocity that the point takes after the rebound

    Dim vDot      As Double
    Dim D         As Double
    Dim NwX       As Double
    Dim NwY       As Double

    '    D = (WallX * WallX + WallY * WallY)
    '    If D = 0 Then Exit Function
    '    D = 1 / Sqr(D)
    D = 1

    NwX = WallX * D
    NwY = WallY * D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = VX * NwX + VY * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    rX = -VX + NwX
    rY = -VY + NwY

End Function
