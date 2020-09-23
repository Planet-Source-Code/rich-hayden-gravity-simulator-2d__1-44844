VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gravity Simulator"
   ClientHeight    =   11520
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15360
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRRandom 
      Caption         =   "Restart with Random Spread"
      Height          =   375
      Left            =   12600
      TabIndex        =   13
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Frame grpControls 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11295
      Left            =   12360
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox picMap 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H000000FF&
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2595
         ScaleWidth      =   2595
         TabIndex        =   14
         Top             =   8520
         Width           =   2655
      End
      Begin VB.Timer tmrMain 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2040
         Top             =   7560
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton cmdRestart 
         Caption         =   "Restart"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Width           =   2415
      End
      Begin MSComctlLib.Slider sldGridSize 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   500
         Min             =   1
         Max             =   100000
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   0
         Value           =   1
      End
      Begin MSComctlLib.Slider sldObjSize 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   0
         Value           =   1
      End
      Begin MSComctlLib.Slider sldRenderSpeed 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   50
         Min             =   1
         Max             =   5000
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   0
         Value           =   1
      End
      Begin MSComctlLib.Slider sldGravity 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   50
         Min             =   1
         Max             =   2000
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   0
         Value           =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Use the below box to move the window around the entire map:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   7800
         Width           =   2415
      End
      Begin VB.Label lblAdd 
         Caption         =   "Click on the render area to add a particle."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Gravity Strength:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Render Speed:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Object Relative Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Window Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   11295
      Left            =   120
      ScaleHeight     =   11235
      ScaleWidth      =   12075
      TabIndex        =   0
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (c) 2003 Richard Hayden. All Rights Reserved.
'
'You may use any code contained in this project for NON-commercial gain
'as long as credit is given to the author (Richard Hayden).
'
'Thanks for downloading my code!

Const dblTotalSize As Double = 100000
Const dblSpreadAreaSize As Double = 100

Dim intNumStars As Integer
Dim dblGridSize As Double
Dim vecWindowCentre As New cls2DVector
Dim dblMassDivider As Double
Dim dblG As Double
Dim dblTime As Double
Dim dblVX As Double
Dim dblVY As Double

Dim blnPause As Boolean
Dim blnDoingVeloc As Boolean

Dim Stars(0 To 99) As clsStar
Dim blnStarCollisions(0 To 99, 0 To 99) As Boolean

Private Sub cmdPause_Click()
    If cmdPause.Caption = "Pause" Then
        blnPause = True
        cmdPause.Caption = "Resume"
    Else
        blnPause = False
        cmdPause.Caption = "Pause"
    End If
End Sub

Private Sub cmdRestart_Click()
    intNumStars = 0
    Form_Load
End Sub

Private Sub cmdRRandom_Click()
    Dim strNum As String
    
    strNum = InputBox("Enter the number of random particles to start with:", "Gravity Simulator")
        
    While Not IsNumeric(strNum)
        MsgBox "Enter a numeric value!", vbOKOnly + vbExclamation, "Gravity Simulator"
        strNum = InputBox("Enter the number of random particles to start with:", "Gravity Simulator")
    Wend
    
    If (CDbl(strNum) > 100) Or (CDbl(strNum) < 1) Then
        MsgBox "Values must be between 1 and 100!", vbOKOnly + vbExclamation, "Gravity Simulator"
        Exit Sub
    End If
    
    intNumStars = CInt(strNum)
    
    Form_Load
End Sub

Private Sub Form_Load()
    Rnd -1
    Randomize
    
    vecWindowCentre.x = 0
    vecWindowCentre.y = 0
    
    dblGridSize = 100
    dblMassDivider = 5000000000#
    dblG = 0.0000000000667
    dblTime = 0.5
    
    picMap.Scale (-dblTotalSize, -dblTotalSize)-(dblTotalSize, dblTotalSize)
    
    InitSliders

    For i = 0 To 99
        Set Stars(i) = New clsStar
    Next i

    For i = 0 To intNumStars - 1
        Stars(i).Position.x = Int((2 * dblSpreadAreaSize * Rnd) - dblSpreadAreaSize)
        Stars(i).Position.y = Int((2 * dblSpreadAreaSize * Rnd) - dblSpreadAreaSize)
        Stars(i).Velocity.x = ((0.2 * Rnd) - 0.1)
        Stars(i).Velocity.y = ((0.2 * Rnd) - 0.1)
        Stars(i).Mass = Int((5000000000# * Rnd) + 500)
    Next i
    
    For i = 0 To 99
        For j = 0 To 99
            blnStarCollisions(i, j) = False
        Next j
    Next i
    
    frmMain.Show
    DoEvents
    MainLoop
End Sub

Private Sub CalculateVariables()
    dblGridSize = sldGridSize.Value
    
    dblMassDivider = 5000000000# / sldObjSize.Value
    
    dblG = 6.67 * 10 ^ ((sldGravity.Value / 100) - 21)
    
    dblTime = 5 * 10 ^ ((sldRenderSpeed.Value / 1000) - 3)
End Sub

Private Sub InitSliders()
    sldGridSize.Value = dblGridSize
    
    sldObjSize.Value = 5000000000# / dblMassDivider
    
    sldGravity.Value = (100 * Log(dblG / 6.67)) / Log(10) + 2100
    
    sldRenderSpeed.Value = (1000 * Log(dblTime / 5)) / Log(10) + 3000
End Sub

Private Sub RenderScene()
    On Error Resume Next
    Dim dblForce As Double
    Dim dblDistBetweenSqrd As Double
    Dim dblAngleTan As Double
    Dim vctForce As New cls2DVector
    Dim vctAccel As New cls2DVector
    Dim dblMassRatio1 As Double
    Dim dblMassRatio2 As Double
    Dim dblMassRatio3 As Double
    Dim vecTempI As New cls2DVector
    Dim vecTempJ As New cls2DVector
    
    picMain.Cls
    picMap.Cls
    picMain.Scale (-dblGridSize, -dblGridSize)-(dblGridSize, dblGridSize)

    For i = 0 To 99
        If (i < intNumStars) And (Not ((Abs(Stars(i).Position.x) > dblTotalSize) Or (Abs(Stars(i).Position.y) > dblTotalSize))) Then
            vctForce.x = 0
            vctForce.y = 0
            For j = 0 To 99
                If j < intNumStars Then
                    'check not same object
                    If (Not (i = j)) And (Not ((Abs(Stars(j).Position.x) > dblTotalSize) Or (Abs(Stars(j).Position.y) > dblTotalSize))) Then
                        'get dist between them
                        dblDistBetweenSqrd = (Stars(j).Position.x - Stars(i).Position.x) ^ 2 + (Stars(j).Position.y - Stars(i).Position.y) ^ 2
                        
                        'get magnitude of the force from this object
                        dblForce = (dblG * Stars(i).Mass * Stars(j).Mass) / dblDistBetweenSqrd
                        
                        'get gradient of force
                        dblAngleTan = (Stars(j).Position.y - Stars(i).Position.y) / (Stars(j).Position.x - Stars(i).Position.x)
                        
                        If Not AreTogether(Stars(i), Stars(j)) Then
                            'need to add or minus depending on their relative positions
                            If (Stars(i).Position.x < Stars(j).Position.x) Then
                                vctForce.x = vctForce.x + dblForce * Cos(Atn(dblAngleTan))
                                vctForce.y = vctForce.y + dblForce * Sin(Atn(dblAngleTan))
                            ElseIf (Stars(i).Position.x > Stars(j).Position.x) Then
                                vctForce.x = vctForce.x - dblForce * Cos(Atn(dblAngleTan))
                                vctForce.y = vctForce.y - dblForce * Sin(Atn(dblAngleTan))
                            End If
                        End If
    
                        If StarsCollide(Stars(i), i, Stars(j), j) Then
                            'preserve old velocities
                            vecTempI.x = Stars(i).Velocity.x
                            vecTempI.y = Stars(i).Velocity.y
                            
                            vecTempJ.x = Stars(j).Velocity.x
                            vecTempJ.y = Stars(j).Velocity.y
                            
                            dblMassRatio1 = (Stars(i).Mass - Stars(j).Mass) / (Stars(i).Mass + Stars(j).Mass)
                            dblMassRatio2 = (2 * Stars(j).Mass) / (Stars(i).Mass + Stars(j).Mass)
                            dblMassRatio3 = (2 * Stars(i).Mass) / (Stars(i).Mass + Stars(j).Mass)
                            
                            'now change veloc for collision
                            Stars(i).Velocity.x = dblMassRatio1 * vecTempI.x + dblMassRatio2 * vecTempJ.x
                            Stars(i).Velocity.y = dblMassRatio1 * vecTempI.y + dblMassRatio2 * vecTempJ.y
                            
                            Stars(j).Velocity.x = dblMassRatio3 * vecTempI.x - dblMassRatio1 * vecTempJ.x
                            Stars(j).Velocity.y = dblMassRatio3 * vecTempI.y - dblMassRatio1 * vecTempJ.y
                        End If
                    End If
                End If
            Next j
            
            'get accel on star
            vctAccel.x = vctForce.x / Stars(i).Mass
            vctAccel.y = vctForce.y / Stars(i).Mass
            
            'now change velocities
            Stars(i).Velocity.x = Stars(i).Velocity.x + vctAccel.x * dblTime
            Stars(i).Velocity.y = Stars(i).Velocity.y + vctAccel.y * dblTime
            
            'now change displacements
            Stars(i).Position.x = Stars(i).Position.x + Stars(i).Velocity.x * dblTime
            Stars(i).Position.y = Stars(i).Position.y + Stars(i).Velocity.y * dblTime
        
            'draw it on both maps
            picMain.Circle (Stars(i).Position.x - vecWindowCentre.x, Stars(i).Position.y - vecWindowCentre.y), (Stars(i).Mass / dblMassDivider)
            picMap.Circle (Stars(i).Position.x, Stars(i).Position.y), Stars(i).Mass / dblMassDivider
        End If
    Next i
    
    'draw red line box around main window
    picMain.Line (-dblTotalSize - vecWindowCentre.x, dblTotalSize - vecWindowCentre.y)-(dblTotalSize - vecWindowCentre.x, dblTotalSize - vecWindowCentre.y), &HFF&
    picMain.Line (-dblTotalSize - vecWindowCentre.x, dblTotalSize - vecWindowCentre.y)-(-dblTotalSize - vecWindowCentre.x, -dblTotalSize - vecWindowCentre.y), &HFF&
    picMain.Line (-dblTotalSize - vecWindowCentre.x, -dblTotalSize - vecWindowCentre.y)-(dblTotalSize - vecWindowCentre.x, -dblTotalSize - vecWindowCentre.y), &HFF&
    picMain.Line (dblTotalSize - vecWindowCentre.x, dblTotalSize - vecWindowCentre.y)-(dblTotalSize - vecWindowCentre.x, -dblTotalSize - vecWindowCentre.y), &HFF&
    
    'draw map box
    picMap.Line (vecWindowCentre.x - dblGridSize, vecWindowCentre.y + dblGridSize)-(vecWindowCentre.x + dblGridSize, vecWindowCentre.y + dblGridSize), &HFFFFFF
    picMap.Line (vecWindowCentre.x - dblGridSize, vecWindowCentre.y - dblGridSize)-(vecWindowCentre.x + dblGridSize, vecWindowCentre.y - dblGridSize), &HFFFFFF
    picMap.Line (vecWindowCentre.x - dblGridSize, vecWindowCentre.y - dblGridSize)-(vecWindowCentre.x - dblGridSize, vecWindowCentre.y + dblGridSize), &HFFFFFF
    picMap.Line (vecWindowCentre.x + dblGridSize, vecWindowCentre.y - dblGridSize)-(vecWindowCentre.x + dblGridSize, vecWindowCentre.y + dblGridSize), &HFFFFFF
End Sub

Private Function AreTogether(Star1 As clsStar, Star2 As clsStar) As Boolean
    On Error Resume Next
    Dim dblCentreDist As Double
    Dim dblRadiusTotal As Double
    
    dblCentreDist = Sqr((Star1.Position.x - Star2.Position.x) ^ 2 + (Star1.Position.y - Star2.Position.y) ^ 2)
    dblRadiusTotal = (Star1.Mass + Star2.Mass) / dblMassDivider
    
    If dblCentreDist <= (dblRadiusTotal + Sqr((Star1.Velocity.x + Star2.Velocity.x) ^ 2 + (Star1.Velocity.y + Star2.Velocity.y) ^ 2)) Then
        AreTogether = True
    Else
        AreTogether = False
    End If
End Function

Private Function StarsCollide(Star1 As clsStar, ByVal intI As Integer, Star2 As clsStar, ByVal intJ As Integer) As Boolean
    On Error Resume Next
    Dim dblCentreDist As Double
    Dim dblRadiusTotal As Double
    
    dblCentreDist = Sqr((Star1.Position.x - Star2.Position.x) ^ 2 + (Star1.Position.y - Star2.Position.y) ^ 2)
    dblRadiusTotal = (Star1.Mass + Star2.Mass) / dblMassDivider
    
    If intI < intJ Then
        If blnStarCollisions(intI, intJ) Then
            blnStarCollisions(intI, intJ) = False
            StarsCollide = False
        Else
            If dblCentreDist <= dblRadiusTotal Then
                blnStarCollisions(intI, intJ) = True
                StarsCollide = True
            Else
                blnStarCollisions(intI, intJ) = False
                StarsCollide = False
            End If
        End If
    Else
        If blnStarCollisions(intJ, intI) Then
            blnStarCollisions(intJ, intI) = False
            StarsCollide = False
        Else
            If dblCentreDist <= dblRadiusTotal Then
                blnStarCollisions(intJ, intI) = True
                StarsCollide = True
            Else
                blnStarCollisions(intJ, intI) = False
                StarsCollide = False
            End If
        End If
    End If
End Function

Private Sub Form_Resize()
    picMain.Scale (-dblGridSize, -dblGridSize)-(dblGridSize, dblGridSize)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strMass As String
    
    If blnDoingVeloc Then
        Stars(intNumStars - 1).Velocity.x = (x - dblVX) / 10
        Stars(intNumStars - 1).Velocity.y = (y - dblVY) / 10
        
        strMass = InputBox("Enter the particle's mass:", "Gravity Simulator")
        
        If IsNumeric(strMass) Then
            If Not CDbl(strMass) > 0 Then
                strMass = "n"
            End If
        End If
        
        While Not IsNumeric(strMass)
            MsgBox "Masses must be numeric and greater than 0!", vbOKOnly + vbExclamation, "Gravity Simulator"
            strMass = InputBox("Enter the particle's mass:", "Gravity Simulator")
            
            If IsNumeric(strMass) Then
                If Not CDbl(strMass) > 0 Then
                    strMass = "n"
                End If
            End If
        Wend
        
        Stars(intNumStars - 1).Mass = CDbl(strMass) * 1000000000
        
        blnDoingVeloc = False
        
        If cmdPause.Caption = "Pause" Then
            blnPause = False
        Else
            RenderOnly intNumStars
        End If
        
        lblAdd.Caption = "Click on the render area to add a particle."
    End If
End Sub

Private Sub RenderOnly(intUpto As Integer)
    picMain.Cls
    picMap.Cls

    For i = 0 To (intUpto - 1)
        If Not ((Abs(Stars(i).Position.x) > dblTotalSize) Or (Abs(Stars(i).Position.y) > dblTotalSize)) Then
            picMain.Circle (Stars(i).Position.x - vecWindowCentre.x, Stars(i).Position.y - vecWindowCentre.y), Stars(i).Mass / dblMassDivider
            picMap.Circle (Stars(i).Position.x, Stars(i).Position.y), Stars(i).Mass / dblMassDivider
        End If
    Next i
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blnDoingVeloc Then
        RenderOnly intNumStars - 1
        
        picMain.Line (dblVX, dblVY)-(x, y)
    End If
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If intNumStars = 100 Then
        MsgBox "Maximum of 100 particles have already been added!", vbExclamation + vbOKOnly, "Gravity Simulator"
        
        blnPause = False
        
        Exit Sub
    End If
    
    intNumStars = intNumStars + 1
    
    Stars(intNumStars - 1).Position.x = x + vecWindowCentre.x
    Stars(intNumStars - 1).Position.y = y + vecWindowCentre.y
    
    blnPause = True
    blnDoingVeloc = True
    dblVX = x
    dblVY = y
    
    lblAdd.Caption = "Now use your mouse to indicate the velocity of the new particle, and click to finalise."
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not (((x + dblGridSize) > dblTotalSize) Or ((x - dblGridSize) < -dblTotalSize)) Or (((y + dblGridSize) > dblTotalSize) Or ((y - dblGridSize) < -dblTotalSize)) Then
        vecWindowCentre.x = x
        vecWindowCentre.y = y
    End If
End Sub

Private Sub sldGravity_Scroll()
    Dim lnghWnd As Long

    lnghWnd = SendMessage(sldGravity.hwnd, TBM_GETTOOLTIPS, 0, ByVal 0)
    SendMessage lnghWnd, TTM_ACTIVATE, Abs(False), ByVal 0
End Sub

Private Sub sldGridSize_Scroll()
    Dim lnghWnd As Long

    lnghWnd = SendMessage(sldGridSize.hwnd, TBM_GETTOOLTIPS, 0, ByVal 0)
    SendMessage lnghWnd, TTM_ACTIVATE, Abs(False), ByVal 0
End Sub

Private Sub sldObjSize_Scroll()
    Dim lnghWnd As Long

    lnghWnd = SendMessage(sldObjSize.hwnd, TBM_GETTOOLTIPS, 0, ByVal 0)
    SendMessage lnghWnd, TTM_ACTIVATE, Abs(False), ByVal 0
End Sub

Private Sub sldRenderSpeed_Scroll()
    Dim lnghWnd As Long

    lnghWnd = SendMessage(sldRenderSpeed.hwnd, TBM_GETTOOLTIPS, 0, ByVal 0)
    SendMessage lnghWnd, TTM_ACTIVATE, Abs(False), ByVal 0
End Sub

Private Sub MainLoop()
    tmrMain.Enabled = True
End Sub

Private Sub tmrMain_Timer()
    DoEvents
    If Not blnPause Then
        CalculateVariables
        RenderScene
    End If
End Sub
