VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c3DEasyCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' NEEDS mVectors.bas

' 3D CAMERA by
' reexre - miorsoft - Roberto Mior

' WORKS ONLY with mVectorUP = 0, -1 , 0

' LEFT HAND RULES
' https://upload.wikimedia.org/wikipedia/commons/b/b2/3D_Cartesian_Coodinate_Handedness.jpg



Option Explicit

Private Const Epsilon As Double = 0.001
Private Const Deg2Rad As Double = 1.74532925199433E-02    'Degrees to Radians
Private Const Rad2Deg As Double = 57.2957795130823    'Radians to Degrees

Private mPosition As tVec3        'Camera Position
Private mLookAt   As tVec3        'Camera LookAT
Private mScreenCenter As tVec3    'Center coords of screen
Private camRIGHT  As tVec3        'Cam Matrix
Private camUP     As tVec3
Private camFRONT  As tVec3
Private mVectorUP As tVec3        'Vector UP
Public mNearPlaneDist As Double
Private mFarPlaneDist As Double


Private mPitch    As Double
Private mYaw      As Double
Private mZOOM     As Double


Friend Property Get GetFront() As tVec3
    GetFront = camFRONT
End Property
Friend Property Get GetUP() As tVec3
    GetUP = camUP
End Property
Friend Property Get GetRIGHT() As tVec3
    GetRIGHT = camRIGHT
End Property


Friend Property Let Position(v As tVec3)
    mPosition = v
    UPDATE
End Property

Friend Property Get Position() As tVec3
    Position = mPosition
End Property

Friend Property Let lookat(v As tVec3)
    mLookAt = v
    UPDATE
End Property

Friend Property Get lookat() As tVec3
    lookat = mLookAt
End Property

Friend Sub SetPositionAndLookAt(pos As tVec3, Look As tVec3)
    mPosition = pos
    mLookAt = Look
    UPDATE
End Sub

Friend Property Get Direction() As tVec3
    Direction = camFRONT
End Property


Friend Property Let VectorUP(UP As tVec3)
    mVectorUP = Normalize3(UP)
    UPDATE
End Property

Friend Property Get VectorUP() As tVec3
    VectorUP = mVectorUP
End Property


Friend Sub GetRotation(ByRef Yaw As Double, ByRef Pitch As Double)
    UPDATE True
    Pitch = mPitch
    Yaw = mYaw
End Sub

Friend Property Let Zoom(v As Double)
    mZOOM = v
End Property

Friend Property Get Zoom() As Double
    Zoom = mZOOM
End Property

Friend Property Let FarPlane(v As Double)
    mFarPlaneDist = v
End Property

Friend Property Get FarPlane() As Double
    FarPlane = mFarPlaneDist
End Property

Friend Property Let NearPlane(v As Double)
    mNearPlaneDist = v
End Property

Friend Property Get NearPlane() As Double
    NearPlane = mNearPlaneDist
End Property

Friend Sub INIT(CameraFrom As tVec3, CameraTo As tVec3, ScreenCenter As tVec3, UP As tVec3)
    mPosition = CameraFrom
    mLookAt = CameraTo
    mScreenCenter = ScreenCenter
    mVectorUP = Normalize3(UP)

    mNearPlaneDist = 5
    mFarPlaneDist = 1E+32

    mZOOM = 1

    UPDATE
End Sub


Private Sub UPDATE(Optional UpdatePitchYawValues As Boolean = False)
    ' Call this every time you change Camera Position or Target !!!
    '    // camera matrix
    Dim d         As tVec3

    ' Left Hand
    '    camFRONT = Normalize3(DIFF3(mLookAt, mPosition))
    '    camRIGHT = Normalize3(CROSS3(mVectorUP, camFRONT))
    '    camUP = Normalize3(CROSS3(camRIGHT, camFRONT))


    ' Right Hand
    camFRONT = Normalize3(DIFF3(mLookAt, mPosition))
    camRIGHT = Normalize3(CROSS3(camFRONT, mVectorUP))
    camUP = Normalize3(CROSS3(camFRONT, camRIGHT))


    'Debug.Print ToString3(camFRONT)
    'Debug.Print ToString3(camRIGHT)
    'Debug.Print ToString3(camUP)




    If UpdatePitchYawValues Then  '  (GetRotation)
        ' ############    CAMERA mVectorUP  MUST BE 0,-1,0
        'https://math.stackexchange.com/questions/470112/calculate-camera-pitch-yaw-to-face-point
        'D = DIFF3(Position, lookat)

        d = camFRONT              'opposite sign than above
        With d
            mPitch = (-Atan2(Sqr(.X * .X + .Z * .Z), .Y)) * Rad2Deg
            mYaw = (-Atan2(-.X, -.Z) + PIh) * Rad2Deg
        End With

    End If

End Sub



'******************************************************************
'   TODO:
'   http://paulbourke.net/geometry/rotate/
'   http://paulbourke.net/geometry/rotate/source.c
'******************************************************************
Public Sub SetRotation(ByVal Yaw As Double, ByVal Pitch As Double)

    Dim d         As Double
    ' Thanks to Passel:
    ' http://www.vbforums.com/showthread.php?870755-3D-Swimming-Fish-Algorithm&p=5356667&viewfull=1#post5356667

    ' ############    CAMERA mVectorUP MUST BE 0,-1,0
    d = Length3(DIFF3(mPosition, mLookAt))
    mPosition.X = mLookAt.X + d * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    mPosition.Y = mLookAt.Y + d * (Sin(Pitch * Deg2Rad))
    mPosition.Z = mLookAt.Z + d * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))

    '    '    'cameraUP = Z
    '            D = Length3(DIFF3(.mPosition, .mLookAt))
    '            mPosition.X = mLookAt.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '            mPosition.Y = mLookAt.Y + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    '            mPosition.Z = mLookAt.Z + D * (Sin(Pitch * Deg2Rad))

    UPDATE

End Sub

Friend Function PointToScreenWDCam(WorldPos As tVec3, ProjectedDistFromCam As Double) As tVec3
    'Point to Screen with Distance^2 From Cam
    Dim p         As tVec3
    Dim S         As tVec3
    Dim IZ        As Double

    S = DIFF3(WorldPos, mPosition)

    p.X = DOT3(S, camRIGHT)
    p.Y = DOT3(S, camUP)
    p.Z = DOT3(S, camFRONT)


    IZ = 1# / p.Z * mZOOM
    PointToScreenWDCam.X = p.X * IZ * mScreenCenter.X + mScreenCenter.X
    PointToScreenWDCam.Y = p.Y * IZ * mScreenCenter.X + mScreenCenter.Y

    PointToScreenWDCam.Z = IZ     ' if its negative point is behind camera
    ProjectedDistFromCam = p.Z    ' if its negative point is behind camera

End Function

Friend Sub PointToScreenCoords(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
                               rX As Double, rY As Double, rZ As Double, InvZ As Double)
    Dim S         As tVec3
    Dim p         As tVec3
    '    Dim InvZ        As Double

    S = vec3(X, Y, Z)
    S = DIFF3(S, mPosition)

    p.X = DOT3(S, camRIGHT)
    p.Y = DOT3(S, camUP)
    p.Z = DOT3(S, camFRONT)

    InvZ = 1 / p.Z
    rX = p.X * InvZ * mZOOM * mScreenCenter.X + mScreenCenter.X
    rY = p.Y * InvZ * mZOOM * mScreenCenter.X + mScreenCenter.Y
    rZ = p.Z                      ' if its negative point is behind camera

End Sub

Friend Sub LineToScreen(P1 As tVec3, P2 As tVec3, Ret1 As tVec3, Ret2 As tVec3, visible As Boolean)

    Dim PlaneCenter As tVec3
    Dim PlaneNormal As tVec3
    Dim IntersectP1 As tVec3
    Dim IntersectP2 As tVec3
    Dim DfromCam1 As Double
    Dim DfromCam2 As Double

    Ret1 = PointToScreenWDCam(P1, DfromCam1)
    Ret2 = PointToScreenWDCam(P2, DfromCam2)

    visible = False


    If DfromCam1 < mNearPlaneDist Then

        If DfromCam2 < mNearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P1 Behind, So Find it's intersection To Near plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mNearPlaneDist))
        IntersectP1 = RayPlaneIntersect(DIFF3(P2, P1), P1, PlaneNormal, PlaneCenter)
        Ret1 = PointToScreenWDCam(IntersectP1, DfromCam1)
    ElseIf DfromCam2 < mNearPlaneDist Then

        If DfromCam1 < mNearPlaneDist Then Exit Sub    'Both points behind camera so EXIT

        'Just P2 Behind, So Find it's intersection To Near plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mNearPlaneDist))
        IntersectP2 = RayPlaneIntersect(DIFF3(P1, P2), P2, PlaneNormal, PlaneCenter)
        Ret2 = PointToScreenWDCam(IntersectP2, DfromCam2)
    End If


    'Same for FAR PLANE
    If DfromCam1 > mFarPlaneDist Then

        If DfromCam2 > mFarPlaneDist Then Exit Sub    'Both points Too far from camera so EXIT

        'Just P1 Too far from, So Find it's intersection To Far plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mFarPlaneDist))
        IntersectP1 = RayPlaneIntersect(DIFF3(P2, P1), P1, PlaneNormal, PlaneCenter)
        Ret1 = PointToScreenWDCam(IntersectP1, DfromCam1)
    ElseIf DfromCam2 > mFarPlaneDist Then

        If DfromCam1 > mFarPlaneDist Then Exit Sub    'Both points Too far from camera so EXIT

        'Just P2 Too far from, So Find it's intersection To Far plane
        PlaneNormal = camFRONT
        PlaneCenter = SUM3(mPosition, MUL3(PlaneNormal, mFarPlaneDist))
        IntersectP2 = RayPlaneIntersect(DIFF3(P1, P2), P2, PlaneNormal, PlaneCenter)
        Ret2 = PointToScreenWDCam(IntersectP2, DfromCam2)
    End If

    visible = True

End Sub


Friend Function IsPointVisibleCoords(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Boolean

    If Z < 0 Then Exit Function   ' behind

    If Z > mFarPlaneDist Then Exit Function    ' too far

    If X < 0 Then Exit Function   ' check if outside screen
    If Y < 0 Then Exit Function
    If X > mScreenCenter.X * 2 Then Exit Function
    If Y > mScreenCenter.Y * 2 Then Exit Function

    IsPointVisibleCoords = True

End Function


Friend Function IsPointVisible(v As tVec3) As Boolean

    With v
        If .Z < 0 Then Exit Function    ' behind
        If .Z > mFarPlaneDist Then Exit Function    ' too far

        If .X < 0 Then Exit Function    ' check if outside screen
        If .Y < 0 Then Exit Function
        If .X > mScreenCenter.X * 2 Then Exit Function
        If .Y > mScreenCenter.Y * 2 Then Exit Function
    End With

    IsPointVisible = True

End Function

Friend Function IsPointVisibleGap(v As tVec3, GAP As Double) As Boolean

    With v
        If .Z < 0 Then Exit Function    ' behind
        If .Z - GAP > mFarPlaneDist Then Exit Function    ' too far

        If .X < -GAP Then Exit Function    ' check if outside screen
        If .Y < -GAP Then Exit Function
        If .X - GAP > mScreenCenter.X * 2 Then Exit Function
        If .Y - GAP > mScreenCenter.Y * 2 Then Exit Function
    End With

    IsPointVisibleGap = True

End Function

Friend Function FOLLOW(Target As tVec3, LookAtSpeed01 As Double, PositionSpeed01 As Double, MAXDist2 As Double, Optional MINDist2 As Double = 0, Optional PosMinY As Double = 0)
    Dim d         As Double

    mLookAt = SUM3(MUL3(mLookAt, 1 - LookAtSpeed01), MUL3(Target, LookAtSpeed01))

    d = Length32(DIFF3(mLookAt, mPosition))
    If d > MAXDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camFRONT, -Sqr(MAXDist2))), PositionSpeed01))
    ElseIf d < MINDist2 Then
        mPosition = SUM3(MUL3(mPosition, 1 - PositionSpeed01), MUL3(SUM3(Target, MUL3(camFRONT, -Sqr(MINDist2))), PositionSpeed01))
    End If

    If PosMinY Then If mPosition.Y > PosMinY Then mPosition.Y = PosMinY

    UPDATE

End Function

Friend Function DistFromCamera2(v As tVec3) As Double
    Dim dx        As Double
    Dim dy        As Double
    Dim Dz        As Double
    dx = v.X - mPosition.X
    dy = v.Y - mPosition.Y
    Dz = v.Z - mPosition.Z
    DistFromCamera2 = dx * dx + dy * dy + Dz * Dz
End Function


