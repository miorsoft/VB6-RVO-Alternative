Attribute VB_Name = "mLegs"
Option Explicit

Public Type tLegs
    Lfeet         As tVec3
    Rfeet         As tVec3
    LfeetToGO     As tVec3
    RfeetToGO     As tVec3

    FrontPhase    As Double

    FF            As Double

End Type

Public FEETs()    As tLegs

Public FrontDist#


'This is a modified fMOD function; it handles negative values specially to ensure they work with certain distort functions
Private Function fMOD(quotient As Double, divisor As Double) As Double
'If quotient < 0 Then Stop

    fMOD = quotient - Fix(quotient / divisor) * divisor
    '    fMOD = quotient - Fix(Abs(quotient / divisor)) * Abs(divisor)
    If (fMOD < 0#) Then fMOD = fMOD + divisor
    '    If (quotient < 0#) Then fMOD = -fMOD

End Function

Public Sub DoFeets()
    Dim I         As Long


    For I = 1 To NA

        With FEETs(I)
            .LfeetToGO.Z = 0
            .LfeetToGO.X = Agent(I).X + Agent(I).VX * 2    'Agent(I).V * 2
            .LfeetToGO.Y = Agent(I).Y + Agent(I).VY * 2

            .RfeetToGO = .LfeetToGO

            FrontDist = Agent(I).Walked

            .FrontPhase = Cos(FrontDist * PI / 3)

            .FF = fMOD(FrontDist / 3, 1#)    '* Agent(I).maxV * 0.25

            If .FrontPhase < 0# Then .FF = -.FF

            If .FF > 0 Then
                .FF = .FF * .FF * (3# - 2# * .FF)    'smoothstep 2023
            Else
                .FF = -.FF * .FF * (3# + 2 * .FF)
            End If


            '            If I = 1 Then Stop


            If .FrontPhase > 0# Then
                .Lfeet.X = .Lfeet.X * (1# - .FF) + .LfeetToGO.X * .FF
                .Lfeet.Y = .Lfeet.Y * (1# - .FF) + .LfeetToGO.Y * .FF

                ' .Rfeet.X = .Rfeet.X - Agent(I).V

            Else
                .Rfeet.X = .Rfeet.X * (1# + .FF) + .RfeetToGO.X * -.FF
                .Rfeet.Y = .Rfeet.Y * (1# + .FF) + .RfeetToGO.Y * -.FF

                ' .Lfeet.X = .Lfeet.X - Agent(I).V
            End If


            '            If I = 1 Then
            '                Debug.Print FrontPhase, FF, .Lfeet.X, .Rfeet.X
            '                Stop
            '            End If

        End With

    Next


End Sub
