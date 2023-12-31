VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpatialGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type tCELL
    pIDX()        As Long
    NP            As Long
    MaxNP         As Long
End Type


Private mW        As Long
Private mH        As Long
Private GridSize  As Long

Private CellW     As Long
Private CellH     As Long


Private CELL()    As tCELL
Private NP        As Long
Private MaxNP     As Long
Private pX()      As Double
Private pY()      As Double

Private MaxDistance2 As Double


Private PairP1()  As Long
Private PairP2()  As Long
Private PairDX()  As Double
Private PairDY()  As Double
Private PairD()   As Double
Private Npairs    As Long
Private MAXNpair  As Long


'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)

Private Declare Function GetMem4 Lib "MSVBVM60" (ByVal Src As Long, ByVal dest As Long) As Long
Private Declare Function ArrPtr Lib "MSVBVM60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Function vbaCopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, Dst As Long, Src As Long) As Long




Public Property Get NofPoints() As Long
    NofPoints = NP
End Property

Public Property Get getPointX(I&) As Double
    getPointX = pX(I)
End Property
Public Property Get getPointY(I&) As Double
    getPointY = pY(I)
End Property

Public Property Get getWorldWidth() As Long
    getWorldWidth = mW
End Property

Public Property Get getWorldHeight() As Long
    getWorldHeight = mH
End Property
Public Property Get getMaxDist() As Long
    getMaxDist = GridSize
End Property

Private Sub pvArrCopySingle(dest() As Double, Src() As Double)
    Dim Size      As Long
    Dim w         As Long
    Dim h         As Long
    '    W = UBound(Src, 1)
    '    H = UBound(Src, 2)
    '    Size = (W + 1) * (H + 1) * LenB(Src(0, 0))
    If pvArrayExists(ArrPtr(Src)) Then

        w = UBound(Src)
        Size = (w + 1) * LenB(Src(0))

        If pvArrayExists(ArrPtr(dest)) Then
            '        If (W - UBound(dest, 1)) Or (H - UBound(dest, 2)) Then
            If (w - UBound(dest, 1)) Then
                'ReDim dest(W, H)
                ReDim dest(w)
            End If
        Else
            '''Array DEST has No Dimension
            'ReDim dest(W, H)
            ReDim dest(w)
        End If

        '        '    CopyMemory ByVal VarPtr(dest(0, 0)), ByVal VarPtr(Src(0, 0)), Size
        '
        '        CopyMemory ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0)), Size


        vbaCopyBytes Size, ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0))
    End If

End Sub

Private Sub pvArrCopyLong(dest() As Long, Src() As Long)
    Dim Size      As Long
    Dim w         As Long
    Dim h         As Long
    '    W = UBound(Src, 1)
    '    H = UBound(Src, 2)
    '    Size = (W + 1) * (H + 1) * LenB(Src(0, 0))

    If pvArrayExists(ArrPtr(Src)) Then

        w = UBound(Src)
        Size = (w + 1) * LenB(Src(0))

        If pvArrayExists(ArrPtr(dest)) Then
            '        If (W - UBound(dest, 1)) Or (H - UBound(dest, 2)) Then
            If (w - UBound(dest, 1)) Then
                'ReDim dest(W, H)
                ReDim dest(w)
            End If
        Else
            '''Array DEST has No Dimension
            'ReDim dest(W, H)
            ReDim dest(w)
        End If

        '    CopyMemory ByVal VarPtr(dest(0, 0)), ByVal VarPtr(Src(0, 0)), Size
        ' CopyMemory ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0)), Size

        vbaCopyBytes Size, ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0))


    End If

End Sub
Private Function pvArrayExists(ByVal ppArray As Long) As Long
    GetMem4 ppArray, VarPtr(pvArrayExists)
End Function


Friend Sub INIT(WorldW As Long, WorldH As Long, MaxDist As Long)
    mW = WorldW
    mH = WorldH

    GridSize = MaxDist
    MaxDistance2 = MaxDist * MaxDist

    If CellW <> mW \ GridSize Or CellH <> mH \ GridSize Then    'Per Mantenere MaxNP
        CellW = mW \ GridSize
        CellH = mH \ GridSize
        ReDim CELL(CellW, CellH)
    End If

    NP = 0
    Npairs = 0

End Sub


Friend Sub ResetPoints()

    Dim X         As Long
    Dim Y         As Long

    For Y = 0 To CellH
        For X = 0 To CellW
            CELL(X, Y).NP = 0&
        Next
    Next
    NP = 0
    '        MaxNP = 0

End Sub
Friend Sub InsertPoint(ByVal X As Double, ByVal Y As Double)

    NP = NP + 1
    If NP > MaxNP Then
        MaxNP = NP * 2
        'If NP > UBound(pX) Then
        MaxNP = NP * 3
        ReDim Preserve pX(MaxNP)
        ReDim Preserve pY(MaxNP)
    End If
    pX(NP) = X
    pY(NP) = Y

    pvAddToCell X \ GridSize, Y \ GridSize, NP

End Sub
'''Friend Sub InsertPointWithIDX(ByVal X As Double, ByVal Y As Double, IDX As Long)
'''
'''
'''    NP = NP + 1
'''    'If NP > MaxNP Then
'''    '    MaxNP = NP * 2
'''    If NP > UBound(pX) Then
'''        MaxNP = NP * 3
'''        ReDim Preserve pX(MaxNP)
'''        ReDim Preserve pY(MaxNP)
'''    End If
'''    pX(NP) = X
'''    pY(NP) = Y
'''
'''    pvAddToCell X \ GridSize, Y \ GridSize, IDX 'NP
'''
'''End Sub
Friend Sub InsertALLpoints(X() As Double, Y() As Double, Optional LimitNofPointTo As Long = 0)
    'Remember to Call ResetPoints First

    Dim I         As Long
    Dim U         As Long



    If LimitNofPointTo Then       ' To avoid pre-Call REDIMing of X Y in some cases
        U = LimitNofPointTo
    Else
        U = UBound(X)
    End If

    For I = 1 To U
        NP = NP + 1
        If NP > MaxNP Then


            '            MaxNP = NP + 64    '* 2
            MaxNP = 1 + NP * 1.5
            '            ReDim Preserve PX(MaxNP)
            '            ReDim Preserve PY(MaxNP)
        End If
        '        PX(NP) = X(I)
        '        PY(NP) = Y(I)
        pvAddToCell X(I) \ GridSize, Y(I) \ GridSize, NP
    Next
    pvArrCopySingle pX, X
    pvArrCopySingle pY, Y

End Sub

Private Sub pvAddToCell(cellX As Long, cellY As Long, currPindex As Long)


    With CELL(cellX, cellY)
        .NP = .NP + 1
        If .NP > .MaxNP Then
            .MaxNP = .NP + 64     '* 2
            ReDim Preserve .pIDX(.MaxNP)
        End If
        .pIDX(.NP) = currPindex
    End With

End Sub


Friend Sub GetPairs(RP1() As Long, RP2() As Long, Rdx() As Double, Rdy() As Double, rPairsCount As Long)

    'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim dx        As Double
    Dim dy        As Double
    Dim I         As Long
    Dim J         As Long

    Dim cWm1      As Long
    Dim cHm1      As Long
    Dim Xp1       As Long
    Dim Yp1       As Long
    Dim Ym1       As Long



    Dim IJ        As Long
    Dim ii        As Long

    Dim PXI       As Double
    Dim PYI       As Double



    cWm1 = CellW - 1&
    cHm1 = CellH - 1&

    Npairs = 0&
    '      O
    '     /
    '    X-O
    '    |\
    '    O O


    For Y = 1& To cHm1
        Yp1 = Y + 1&
        Ym1 = Y - 1&
        For X = 0& To cWm1
            Xp1 = X + 1&

            With CELL(X, Y)


                For I = 1& To .NP    '- 1& ' Should be -1& to do only SELF but so we can do even others
                    ii = .pIDX(I)
                    PXI = pX(ii)
                    PYI = pY(ii)
                    For J = I + 1& To .NP    'SELF
                        IJ = .pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                    Next


                    'RIGHT
                    For J = 1& To CELL(Xp1, Y).NP
                        IJ = CELL(Xp1, Y).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                    Next

                    'down
                    For J = 1& To CELL(X, Yp1).NP
                        IJ = CELL(X, Yp1).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                    Next

                    'DownRight Diag
                    For J = 1& To CELL(Xp1, Yp1).NP
                        IJ = CELL(Xp1, Yp1).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                    Next

                    'upRight Diag
                    For J = 1& To CELL(Xp1, Ym1).NP
                        IJ = CELL(Xp1, Ym1).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                    Next
                Next
            End With

        Next
    Next



    '-------- NOW Bottom ROW
    '      O
    '     /
    '    X-O

    Y = CellH
    Ym1 = Y - 1&
    For X = 0& To cWm1
        Xp1 = X + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1& to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = pX(ii)
                PYI = pY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, Y).NP
                    IJ = CELL(Xp1, Y).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
                'upRight Diag
                For J = 1& To CELL(Xp1, Ym1).NP
                    IJ = CELL(Xp1, Ym1).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
            Next

        End With

        '-------- NOW TOP ROW
        '    X-O
        '    |\
        '    O O


        With CELL(X, 0&)
            For I = 1& To .NP     '- 1& ' Should be -1& to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = pX(ii)
                PYI = pY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, 0&).NP
                    IJ = CELL(Xp1, 0&).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
                'DownRight Diag
                For J = 1& To CELL(Xp1, 1&).NP
                    IJ = CELL(Xp1, 1&).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
                'Down
                For J = 1& To CELL(X, 1&).NP
                    IJ = CELL(X, 1&).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
            Next

        End With
    Next X



    ' -----------------------------RIGHT Column
    '    X
    '    |
    '    O
    X = CellW
    For Y = 0& To cHm1
        Yp1 = Y + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1& to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = pX(ii)
                PYI = pY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
                ' DOWN  )
                For J = 1& To CELL(X, Yp1).NP
                    IJ = CELL(X, Yp1).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
                Next
            Next
        End With
    Next


    ' -------- FINALLY bottm right cell
    With CELL(CellW, CellH)
        'SELF
        For I = 1& To .NP - 1&
            ii = .pIDX(I)
            PXI = pX(ii)
            PYI = pY(ii)
            For J = I + 1& To .NP
                IJ = .pIDX(J)
                dx = pX(IJ) - PXI
                dy = pY(IJ) - PYI
                If dx * dx + dy * dy < MaxDistance2 Then pvAddPair ii, IJ, dx, dy
            Next
        Next
    End With




    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs
    '    rP1 = PairP1
    '    rP2 = PairP2
    '    rDX = PairDX
    '    rDY = PairDY

    'FASTER
    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY


    '    'Check for ERRORS (Duplicates)
    '        For X = 1& To Npairs - 1&
    '            For Y = X + 1& To Npairs
    '                If (PairP1(X) = PairP1(Y)) And (PairP2(X) = PairP2(Y)) Then Stop
    '                If (PairP1(X) = PairP2(Y)) And (PairP2(X) = PairP1(Y)) Then Stop
    '            Next
    '        Next

End Sub
Friend Sub GetPairsWDist(RP1() As Long, RP2() As Long, Rdx() As Double, Rdy() As Double, RD() As Double, rPairsCount As Long)

    'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim dx        As Double
    Dim dy        As Double
    Dim I         As Long
    Dim J         As Long
    Dim d         As Double

    Dim cWm1      As Long
    Dim cHm1      As Long
    Dim Xp1       As Long
    Dim Yp1       As Long
    Dim Ym1       As Long


    Dim IJ        As Long
    Dim ii        As Long

    Dim PXI       As Double
    Dim PYI       As Double


    cWm1 = CellW - 1&
    cHm1 = CellH - 1&


    Npairs = 0&
    '      O
    '     /
    '    X-O
    '    |\
    '    O O

    For Y = 1& To cHm1
        Yp1 = Y + 1&
        Ym1 = Y - 1&
        For X = 0& To cWm1
            Xp1 = X + 1&
            With CELL(X, Y)

                For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others
                    ii = .pIDX(I)
                    PXI = pX(ii)
                    PYI = pY(ii)
                    For J = I + 1& To .NP    'SELF
                        IJ = .pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        d = dx * dx + dy * dy
                        If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                    Next

                    'RIGHT
                    For J = 1& To CELL(Xp1, Y).NP
                        IJ = CELL(Xp1, Y).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        d = dx * dx + dy * dy
                        If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d

                    Next

                    'down
                    For J = 1& To CELL(X, Yp1).NP
                        IJ = CELL(X, Yp1).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        d = dx * dx + dy * dy
                        If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d

                    Next

                    'DownRight Diag
                    For J = 1& To CELL(Xp1, Yp1).NP
                        IJ = CELL(Xp1, Yp1).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        d = dx * dx + dy * dy
                        If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d

                    Next

                    'upRight Diag
                    For J = 1& To CELL(Xp1, Ym1).NP
                        IJ = CELL(Xp1, Ym1).pIDX(J)
                        dx = pX(IJ) - PXI
                        dy = pY(IJ) - PYI
                        d = dx * dx + dy * dy
                        If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                    Next
                Next
            End With

        Next
    Next

    '-------- NOW Bottom ROW
    '      O
    '     /
    '    X-O

    Y = CellH
    Ym1 = Y - 1&
    For X = 0& To cWm1
        Xp1 = X + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1 to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = pX(ii)
                PYI = pY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d

                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, Y).NP
                    IJ = CELL(Xp1, Y).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                Next
                'upRight Diag
                For J = 1& To CELL(Xp1, Ym1).NP
                    IJ = CELL(Xp1, Ym1).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                Next
            Next

        End With


        '-------- NOW TOP ROW
        '    X-O
        '    |\
        '    O O
        With CELL(X, 0&)
            For I = 1& To .NP     '- 1& ' Should be -1 to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = pX(ii)
                PYI = pY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, 0&).NP
                    IJ = CELL(Xp1, 0&).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                Next
                'DownRight Diag
                For J = 1& To CELL(Xp1, 1&).NP
                    IJ = CELL(Xp1, 1&).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                Next
                'Down
                For J = 1& To CELL(X, 1&).NP
                    IJ = CELL(X, 1&).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
                Next
            Next

        End With
    Next X



    ' RIGHT Column
    '    X
    '    |
    '    O
    X = CellW
    For Y = 0& To cHm1
        Yp1 = Y + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1 to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = pX(ii)
                PYI = pY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d

                Next
                ' DOWN  )
                For J = 1& To CELL(X, Yp1).NP
                    IJ = CELL(X, Yp1).pIDX(J)
                    dx = pX(IJ) - PXI
                    dy = pY(IJ) - PYI
                    d = dx * dx + dy * dy
                    If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d

                Next
            Next
        End With
    Next


    ' -------- FINALLY bottm right cell
    With CELL(CellW, CellH)
        'SELF
        For I = 1& To .NP - 1&
            ii = .pIDX(I)
            PXI = pX(ii)
            PYI = pY(ii)
            For J = I + 1& To .NP
                IJ = .pIDX(J)
                dx = pX(IJ) - PXI
                dy = pY(IJ) - PYI
                d = dx * dx + dy * dy
                If d < MaxDistance2 Then pvAddPairWDist ii, IJ, dx, dy, d
            Next
        Next
    End With




    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY
    pvArrCopySingle RD, PairD



End Sub

Private Sub pvAddPair(P1 As Long, P2 As Long, dx As Double, dy As Double)
    Npairs = Npairs + 1
    If Npairs > MAXNpair Then
        MAXNpair = Npairs + 128   '* 2
        ReDim Preserve PairP1(MAXNpair)
        ReDim Preserve PairP2(MAXNpair)
        ReDim Preserve PairDX(MAXNpair)
        ReDim Preserve PairDY(MAXNpair)
    End If

    PairP1(Npairs) = P1
    PairP2(Npairs) = P2
    PairDX(Npairs) = dx
    PairDY(Npairs) = dy

End Sub

Private Sub pvAddPairWDist(P1 As Long, P2 As Long, dx As Double, dy As Double, d As Double)

    Npairs = Npairs + 1

    If Npairs > MAXNpair Then
        '        MAXNpair = Npairs + 128    '* 2

        If Npairs > 4 Then
            MAXNpair = Npairs * 1.3
        Else
            MAXNpair = Npairs + 8

        End If

        ReDim Preserve PairP1(MAXNpair)
        ReDim Preserve PairP2(MAXNpair)
        ReDim Preserve PairDX(MAXNpair)
        ReDim Preserve PairDY(MAXNpair)
        ReDim Preserve PairD(MAXNpair)
    End If

    PairP1(Npairs) = P1
    PairP2(Npairs) = P2
    PairDX(Npairs) = dx
    PairDY(Npairs) = dy
    PairD(Npairs) = d

End Sub



Friend Sub QueryPointsInCircle(cX As Double, cY As Double, r As Double, rdist() As Double, rIdx() As Long, UBcount As Long)

    Dim X         As Long
    Dim Y         As Long
    Dim XX        As Long
    Dim YY        As Long

    Dim J         As Long
    Dim IJ        As Long

    Dim dx        As Double
    Dim dy        As Double
    Dim d         As Double
    Dim GG        As Long
    Dim R2        As Double


    Dim XF        As Long
    Dim XT        As Long
    Dim YF        As Long
    Dim YT        As Long

    X = cX \ GridSize
    Y = cY \ GridSize


    ReDim rIdx(0)
    Npairs = 0

    GG = r \ GridSize
    If GG < 1 Then GG = 1         'Commented 2023
    'If GG < 1 Then Stop
    'If GG > 2 Then Debug.Print GG & "  " & R

    R2 = r * r

    XF = X - GG: If XF < 0& Then XF = 0&
    XT = X + GG: If XT > CellW Then XT = CellW
    YF = Y - GG: If YF < 0& Then YF = 0&
    YT = Y + GG: If YT > CellH Then YT = CellH

    For YY = YF To YT
        For XX = XF To XT

            With CELL(XX, YY)
                For J = 1 To .NP
                    IJ = .pIDX(J)
                    dx = pX(IJ) - cX
                    dy = pY(IJ) - cY
                    d = dx * dx + dy * dy
                    If d <= R2 Then
                        pvAddPointWDist IJ, d
                    End If
                Next
            End With
        Next
    Next

    pvArrCopySingle rdist, PairD
    pvArrCopyLong rIdx, PairP1
    UBcount = Npairs




End Sub


Private Sub pvAddPointWDist(P1 As Long, d As Double)

    Npairs = Npairs + 1

    If Npairs > MAXNpair Then

        If Npairs > 4 Then
            MAXNpair = Npairs * 1.3
        Else

            MAXNpair = Npairs * 1.5    'Npairs + 8
        End If

        ReDim Preserve PairP1(MAXNpair)
        'ReDim Preserve PairP2(MAXNpair)
        'ReDim Preserve PairDX(MAXNpair)
        'ReDim Preserve PairDY(MAXNpair)
        ReDim Preserve PairD(MAXNpair)
    End If

    PairP1(Npairs) = P1
    'PairP2(Npairs) = P2
    'PairDX(Npairs) = dx
    'PairDY(Npairs) = dy
    PairD(Npairs) = d

End Sub
