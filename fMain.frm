VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "fMain"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   729
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   988
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Label labelClick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'F
'MsgBox KeyCode
    If KeyCode = 70 Then doFRAMES = Not (doFRAMES)


    'C
    If KeyCode = 67 Then CameraMode = Not (CameraMode)

    If KeyCode = 37 Then FOLLOW = FOLLOW - 1: If FOLLOW < 1 Then FOLLOW = NA
    If KeyCode = 39 Then FOLLOW = FOLLOW + 1: If FOLLOW > NA Then FOLLOW = 1

End Sub

Private Sub Form_Load()
    Randomize Timer

    ScaleMode = vbPixels
    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*", vbArchive) <> vbNullString Then Kill App.Path & "\Frames\*.*"


'    '    ' UNCOMMENT For FRAMES
'    Me.ScaleHeight = 576 - 12     ' 640
'    Me.ScaleWidth = 1024 - 36     '852 ' Round(Me.ScaleHeight * 4 / 3)



    WorldW = 1200                 'Me.ScaleWidth
    WorldH = 1200                 '730 ' Me.ScaleHeight

    Init_RVO 400 * 1.3

    ScreenW = Me.ScaleWidth
    ScreenH = Me.ScaleHeight

    Set CAM = New c3DEasyCam

    CAM.INIT vec3(200, 200, 200), vec3(0, 0, 0), vec3(ScreenW * 0.5, ScreenH * 0.5, 0), vec3(0, -1, 0)
    CAM.NearPlane = 10
    CAM.FarPlane = 3000



    labelClick.Left = (Me.ScaleWidth - labelClick.Width) * 0.5

    RunningNotInIDE = (App.LogMode <> 0)

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    labelClick.visible = False

    If Not (doLOOP) Then Call MAINLOOP
    If Button = 1 Then
        MODE = (MODE + 1) Mod 7
        If MODE = 6 Then
            WorldW = 2000
            WorldH = 2000
            GRID.INIT WorldW * 1, WorldH * 1, GridSize

            INIT_Targets

        Else
            '            WorldW = 1000         'Me.ScaleWidth
            '            WorldH = 730          ' Me.ScaleHeight
            '            GRID.INIT WorldW * 1, WorldH * 1, GridSize
        End If

    End If

    If Button = 2 Then
        If FOLLOW = 0 Then FOLLOWworst: Exit Sub
        If FOLLOW > 1 Then FOLLOW = 1: Exit Sub
        If FOLLOW = 1 Then FOLLOW = 0
    End If
End Sub



Private Sub Form_Resize()

    ScreenW = Int(Me.ScaleWidth)
    ScreenH = Int(Me.ScaleHeight)
    Set CAM = New c3DEasyCam
    CAM.INIT vec3(200, 200, 200), vec3(0, 0, 0), vec3(ScreenW * 0.5, ScreenH * 0.5, 0), vec3(0, -1, 0)
    CAM.NearPlane = 10
    CAM.FarPlane = 3000


    Set SRF = Cairo.CreateSurface(ScreenW, ScreenH, ImageSurface)
    Set CC = SRF.CreateContext
    CC.AntiAlias = CAIRO_ANTIALIAS_FAST
    CC.SetLineCap CAIRO_LINE_CAP_ROUND

End Sub

'Private Sub Form_Unload(Cancel As Integer)
'Stop
'    doloop = False
'
'End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    doLOOP = False

End Sub
