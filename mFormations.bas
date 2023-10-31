Attribute VB_Name = "mFormations"
Option Explicit

Public CurrFormation As Long
Public NFM        As Long

Public Sub INIT_TFormation()
    Dim tW        As Long
    Dim TH        As Long
    Dim I         As Long
    Dim XX        As Long
    Dim YY        As Long
    Dim d         As Double
    Dim r         As Double
    Dim a         As Double
    Dim GAP       As Double

    GAP = ControlDist * 0.7


    NFM = 1

    NTargets = NA
    ReDim ArrayOfTargets(NTargets)

    tW = Sqr(NA)

    Select Case CurrFormation
    Case 0
        For I = 1 To NA
            With Agent(I)
                XX = (I Mod tW)
                YY = (I \ tW)
                ArrayOfTargets(I).X = (XX - tW * 0.5)
                ArrayOfTargets(I).Y = (YY - tW * 0.5)

                .TX = WorldW * 0.5 + (XX - tW * 0.5) * 1.1 * (GAP)    ' WorldW * 0.9 / tW
                .TY = WorldH * 0.5 + (YY - tW * 0.5) * 1.1 * (GAP)    ' WorldW * 0.9 / tW


                d = (XX - tW * 0.5) * (XX - tW * 0.5) + (YY - tW * 0.5) * (YY - tW * 0.5)
                d = Sqr(d)

                .cR = 0.65 + 0.35 * Cos(d * 1.25)
                .cG = 0.65 + 0.35 * Cos(d * 0.5)
                .cB = 0.65 + 0.35 * Sin(d * 1.1)
            End With

        Next
    Case 1

        Dim KA    As Double
        r = GAP
        a = 0
        KA = 8 * PI / r
        For I = 1 To NA
            With Agent(I)

                XX = r * 0.67 * 1.25
                YY = r * 0.21

                .TX = WorldW * 0.5 + Cos(a) * r
                .TY = WorldH * 0.5 + Sin(a) * r
                a = a + KA
                If a + KA * 0.5 > PI2 Then
                    a = a - PI2
                    r = r + GAP
                    KA = 8 * PI / r
                End If
                'Debug.Print R, A, KA
                '


                ArrayOfTargets(I).X = (XX - tW * 0.5)
                ArrayOfTargets(I).Y = (YY - tW * 0.5)
                d = (XX - tW * 0.5) * (XX - tW * 0.5) + (YY - tW * 0.5) * (YY - tW * 0.5)
                d = Sqr(d)
                .cR = 0.65 + 0.35 * Cos(d * 1.25)
                .cG = 0.65 + 0.35 * Cos(d * 0.5)
                .cB = 0.65 + 0.35 * Sin(d * 1.1)
            End With

        Next




    End Select

    CurrFormation = CurrFormation + 1

    'FormationsOptimize


End Sub



Public Sub FormationsOptimize()
    Dim oGW&, oGH&, oGD&
    Dim curGS     As Long

    Dim P1()      As Long
    Dim P2()      As Long
    Dim dx() As Double, dy() As Double, DD() As Double
    Dim Npair     As Long
    Dim I&, J&, K&, PA&

    oGW = GRID.getWorldWidth
    oGH = GRID.getWorldHeight
    oGD = GRID.getMaxDist
    '------------------------------------------------------
    '------------------------------------------------------
    '------------------------------------------------------

    curGS = oGD

    ReDim Preserve Agent(NA * 2)
    For I = 0 To NA * 2
        Agent(I).fTargTo = 0
        Agent(I).fAgentFrom = 0

        Agent(I).fMinD = 1E+32
    Next

Again:
    GRID.ResetPoints
    GRID.INIT oGW, oGH, curGS



    For I = 1 To NA
        GRID.InsertPoint Agent(I).X, Agent(I).Y
    Next




    For I = 1 To NA
        GRID.InsertPoint Agent(I).TX, Agent(I).TY
    Next
    GRID.GetPairsWDist P1(), P2(), dx(), dy(), DD(), Npair


    For PA = 1 To Npair
        I = P1(PA)
        J = P2(PA)
        If Abs(I - J) > NA - 1 Then

            If I > J Then K = I: I = J: J = K

            If DD(PA) < Agent(I).fMinD Then

                If DD(PA) < Agent(Agent(J).fAgentFrom).fMinD Then

                    Agent(Agent(I).fTargTo).fMinD = 1E+32
                    Agent(Agent(J).fAgentFrom).fMinD = 1E+32

                    Agent(Agent(I).fTargTo).fAgentFrom = 0
                    Agent(Agent(I).fTargTo).fAgentFrom = 0

                    Debug.Print I, J, , Agent(I).fTargTo, Agent(I).fTargTo

                    Agent(I).fMinD = DD(PA)
                    Agent(J).fMinD = DD(PA)

                    Agent(I).fTargTo = J
                    Agent(J).fAgentFrom = I
                End If
            End If

        End If
    Next

    K = 0
    For I = 1 To NA
        If Agent(I).fTargTo = 0 Then K = K + 1
    Next
    Debug.Print K, curGS

    curGS = curGS * 1.2
    If K Then If curGS < 4000 Then GoTo Again


    For I = 1 To NA
        If Agent(I).fTargTo Then
            J = Agent(I).fTargTo - NA
            Agent(I).TX = Agent(J).TX
            Agent(I).TY = Agent(J).TY
        End If


    Next
















    '------------------------------------------------------
    '------------------------------------------------------
    '------------------------------------------------------
    'RestoreGRID
    GRID.INIT oGW, oGH, oGD



End Sub
