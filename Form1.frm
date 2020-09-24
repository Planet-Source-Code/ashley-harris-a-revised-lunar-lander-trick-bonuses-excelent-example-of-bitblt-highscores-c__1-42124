VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lunar Lander"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   9000
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8415
      Top             =   1305
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   9270
      Pattern         =   "*.bmp"
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'VB's trig is all based in radians, which, to convert, neto know pi.
'I know nothing about radians, except for the fact that to convert you need
'to know 2pi/360
Private Const pi As Double = 3.141592653559

'Information for the movement and posistion of the player
Public fuel As Long
Public xpos As Single
Public ypos As Single
Public xspeed As Single
Public yspeed As Single
Public angle As Single
Public bonus As Long
Public score As Currency 'long isn't long enough
Public lives As Long
Public playing As Boolean

'Information for the movement and posistion of the landing pad
Public padxpos As Long
Public padxspeed As Single
Public padypos As Long
Public padyspeed As Single
Public padfallout As Boolean

'tells us if the user has done the first half of a super trick
'(land after a sumersalt without thrusting)
Public supertrick As Boolean

'partical system for tracking the portions of the explosion
Private explosionttl As Long
Private expxpos(100) As Single
Private expypos(100) As Single
Private expxspeed(100) As Single
Private expyspeed(100) As Single
Private expcolour(100) As Long

'you can only land upside down once per level
Private upsidedownlandingreqrded As Boolean

'used for working out the frame rate
Private fps As Long
Private fpscount As Long

'set to true when the game wants to end
Public bunload As Boolean

'the force of the wind
Private wind As Single

'what level your on
Public level As Long

Public Sub explode()
    'create an explosion at the point of the player, keeping the average partical speed
    '(as a vector) equal to the speed of the rocket (as a vector)
    explosionttl = 300
    For a = 0 To 100
        expxpos(a) = xpos
        expypos(a) = ypos - 25
        expxspeed(a) = xspeed + Rnd() * 5 - 2.5
        expyspeed(a) = yspeed + Rnd() * 5 - 2.5
        expcolour(a) = RGB(200 + Int(Rnd() * 56), Int(Rnd() * 256), 0)
    Next a
End Sub

Private Sub Form_Click()
    'clicking meas start a new game
        Unload Me
    
End Sub

Private Sub Form_Load()
    'resets all variables and starts the game
    newgame
End Sub

Private Sub Form_Queryunload(Cancel As Integer, bunloadMode As Integer)
    'cleanup, bunload all the graphics

    Me.Visible = False
    bunload = True
    Form2.Show
End Sub

Private Sub Timer1_Timer()
    'the main render loop
    If Not Visible Then Exit Sub
    Do
        t = GetTickCount + 30 'frame rate = 1000/this
        'DoEvents
        If Not playing Then
            'if your dead then draw the spaceship offscreen
            xpos = -500
            ypos = -500
            xspeed = 0
            yspeed = 0
            bonus = 0
            fuel = 0
            Caption = "GAME OVER - Click to continue"
        End If
        
        'set the backbuffer to the picture
        BitBlt backbuffer, 0, 0, 640, 480, lib("skypic640x480.bmp"), 0, 0, vbSrcCopy
        
        'gravity
        yspeed = yspeed + 0.2
        
        'if the space key is down, then thrust. making sure we have fuel to thrust
        If GetAsyncKeyState(vbKeySpace) Then
            If fuel > 0 Then
                fuel = fuel - 1
                'we can rotate, so make sure to thrust in the right direction, cause our
                'engine rotates
                xspeed = xspeed + Sin(angle * 2 * pi / 360)
                yspeed = yspeed - Cos(angle * 2 * pi / 360)
                
                supertrick = False
            End If
        End If
        
        'apply speed to the position
        xpos = xpos + xspeed
        ypos = ypos + yspeed
        
        'if the keys are down, then rotate left
        If GetAsyncKeyState(vbKeyLeft) Then
            angle = angle - 11.25
        End If
        If GetAsyncKeyState(vbKeyRight) Then
            angle = angle + 11.25
        End If
        
        'make sure we don't have angles > 360 or < 0
        If angle > 360 Then angle = 11.25
        If angle < 0 Then angle = 348.75
        
        'work out which graphic to display from the angle
        a = angle / 11.25 + 1
        If a < 10 Then a = "0" & a
        
        'draw the right graphic onto the screen, taking care to draw the transparent bits,
        'as defined in the .msk.bmp. Note we draw to the backbuffer, not the screen
        'the backbuffer is a tempory
        BitBlt backbuffer, xpos - 32, ypos - 64, 64, 64, lib("rocket00" & a & ".msk.bmp"), 0, 0, vbMergePaint
        BitBlt backbuffer, xpos - 32, ypos - 64, 64, 64, lib("rocket00" & a & ".bmp"), 0, 0, vbSrcAnd
        
        'draw the pad in the same manner also to the backbuffer
        BitBlt backbuffer, padxpos - 22, padypos - 20, 45, 23, lib("pad.msk.bmp"), 0, 0, vbMergePaint
        BitBlt backbuffer, padxpos - 22, padypos - 20, 45, 23, lib("pad.bmp"), 0, 0, vbSrcAnd
        
        'make the pad bounce
        If padxpos > 600 And padxspeed > -(level / 10 + 1) Then padxspeed = padxspeed - 1
        If padxpos < 40 And padxspeed < (level / 10 + 1) Then padxspeed = padxspeed + 1
        padxpos = padxpos + padxspeed
        padypos = padypos + padyspeed
        
        'draw the lives count
        If lives >= 1 Then
            BitBlt backbuffer, 576, 416, 64, 64, lib("rocket0001.msk.bmp"), 0, 0, vbMergePaint
            BitBlt backbuffer, 576, 416, 64, 64, lib("rocket0001.bmp"), 0, 0, vbSrcAnd
        End If
        If lives >= 2 Then
            BitBlt backbuffer, 544, 416, 64, 64, lib("rocket0001.msk.bmp"), 0, 0, vbMergePaint
            BitBlt backbuffer, 544, 416, 64, 64, lib("rocket0001.bmp"), 0, 0, vbSrcAnd
        End If
        If lives >= 3 Then
            BitBlt backbuffer, 512, 416, 64, 64, lib("rocket0001.msk.bmp"), 0, 0, vbMergePaint
            BitBlt backbuffer, 512, 416, 64, 64, lib("rocket0001.bmp"), 0, 0, vbSrcAnd
        End If
        
        'add a trick bonus
        If angle > 150 And angle < 220 Then
            bonus = bonus + 10
            supertrick = True
        End If
        
        
        'deduct one from the bonus
        If bonus > 0 And Not padfallout Then bonus = bonus - 1
        
        'if you've landed on the pad, then apply gravity to the pad
        If padfallout Then padyspeed = padyspeed + 0.05
        
        'check for landing on the pad
        If xpos - 5 < padxpos + 20 And xpos + 5 > padxpos - 20 Then
        
            If ypos < padypos And ypos > padypos - 30 Then
                'flyby bonus
                If Abs(xspeed) > 10 Then bonus = bonus + (xspeed + (padypos - ypos)) ^ 2
            End If
            
            If ypos - 5 < padypos And ypos > padypos Then
                If angle < 15 Or angle > 340 Then
                    If (padyspeed - yspeed) < 5 And yspeed > 0 Then
                        'you landed
                        
                        If padfallout Then bonus = bonus + 1000
                        If upsidedownlandingreqrded Then bonus = bonus * 2 + 1000000
                        If upsidedownlandingreqrded And supertrick Then bonus = bonus * 5 + 1000000
                        If supertrick = True And Not padfallout Then bonus = bonus * 20 + 50000
                        
                        padfallout = True
                        xspeed = 0
                        yspeed = -1
                        padxspeed = 0
                        
                    Else
                        'you die (too fast)
                        If Not padfallout Then newround True
                    End If
                Else
                    'you die (too much of a lean)
                    If Not padfallout Then newround True
                End If
            Else
                'you die (hit the pad from below or the sides)
                If Abs(ypos - padypos) < 50 And ypos > padypos Then
                    'however, you can land on the bottom of the pad upside down.
                    'this is worth a lot of bonus points
                    If angle > 150 And angle < 230 And (yspeed - padyspeed) > (-5 + upsidedownlandingreqrded * 5) Then
                        If padyspeed >= 0 Then
                            If upsidedownlandingreqrded Then bonus = bonus + 1000000 Else bonus = (bonus + 10000) * 100
                            padyspeed = padyspeed + yspeed - 1
                            If upsidedownlandingreqrded Then padyspeed = padyspeed / 2
                            padfallout = True
                            yspeed = 2
                            upsidedownlandingreqrded = True
                        End If
                    Else
                        newround True
                    End If
                End If
                
                'you get a trick bonus if you can do a loop under the pad
                'while it's falling down
                If ypos > padypos And padfallout Then
                    If upsidedownlandingreqrded Then
                        bonus = bonus + 1000
                    Else
                        bonus = bonus * 1.1 + 500 + level * 50 + lives * 500 + padypos * 10 + (640 - ypos) * 10
                    End If
                End If
            End If
        End If
        
        'make you 'bounce' when you hit the top
        If padfallout And ypos < 30 Then
            ypos = 30
            yspeed = 0
        End If
        
        'if there is an explosion, then, process it
        If explosionttl > 0 Then
            explosionttl = explosionttl - 1
            For a = 0 To 100
                expxpos(a) = expxpos(a) + expxspeed(a)
                expypos(a) = expypos(a) + expyspeed(a)
                expyspeed(a) = expyspeed(a) + 0.1
                
                'draw the partical
                SetPixel backbuffer, expxpos(a), expypos(a), expcolour(a)
            Next a
        End If
        
        'when the pad falls out the bottom, a new round starts
        If padfallout And padypos > 500 Then
            newround
            padfallout = False
            padypos = Int(Rnd() * 400) + 50
            padxpos = 700
            padyspeed = 0
        End If
        
        'displays the status bar
        If playing Then Caption = "Lunar Lander. Fuel: " & Format(fuel, "000") & " score: " & Format(score, "000000") & " Bonus: " & Format(bonus, "000") & " Level: " & level & " FPS:" & fpscount
        
        'if you fall off the bottom of the screen, you die.
        If ypos > 490 Then
            newround True
        End If
        
        'draw the backbuffer to the screen
        BitBlt hdc, 0, 0, 640, 480, backbuffer, 0, 0, vbSrcCopy
        
        'we just drew a frame, so, count it
        fps = fps + 1
        
        'process the wind force on the speed
        xspeed = xspeed + wind
        DoEvents
        
        While t > GetTickCount
           DoEvents
        Wend
        DoEvents
    Loop Until bunload
    
    Form2.updatehighscore score
    Form2.Show
    Unload Form1
    
End Sub

Public Sub newround(Optional die As Boolean)
    'called when a new set of numbers is needed, ie fuel, bonus, etc.
    'when called after a players death, the die parameter is true
    'when called after an advancement in levels, the die parameter is either false or missing
    If Not die Then score = score + (bonus + fuel) * (lives + 1)
    If die Then lives = lives - 1
    If lives = -1 Then playing = False
    
    If die Then
        explode
        xpos = 320
        ypos = 80
        xspeed = 0
        yspeed = 0
        angle = 0
    Else
        level = level + 1
    End If
    
    bonus = 10 * level + 100
    fuel = 300 - level * 5
    wind = (Rnd() - 0.5) / 5 * (level / 5)
    
    upsidedownlandingreqrded = False
End Sub

Public Sub newgame()
    'called to initialise the game
    level = 1
    newround True
    score = 0
    lives = 3
    playing = True
    padypos = 300
End Sub

Private Sub Timer2_Timer()
    'used to calculate the frame rate, this sub is called exactly every 1 second
    fpscount = fps
    fps = 0
End Sub
