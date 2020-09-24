VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lunar Lander"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label score 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label highscore 
      BackStyle       =   0  'Transparent
      Caption         =   "High scores:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lunar Lander"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Form1.Show
    Form1.newgame
    Form1.playing = True
    Form1.bunload = False
End Sub

Private Sub Form_Load()
    Set Picture = LoadPicture(fso.BuildPath(App.Path, "skypic640x480.bmp"))
    loadhighscores
    
    'loads the backbuferr
    setupbackbuffer
    
    'loads all graphics
    Form1.File1.Path = App.Path
    For a = 0 To Form1.File1.ListCount
        lib(Form1.File1.List(a)) = LoadGraphicDC(fso.BuildPath(App.Path, Form1.File1.List(a)))
    Next a
    
    
End Sub

Public Sub loadhighscores()
    'load all the high scores to the screen
    'creating objects to display them as needed

    Dim ts As TextStream
    Set ts = fso.OpenTextFile(fso.BuildPath(App.Path, "highscores.txt"), ForReading, True)
    c = 1
more:
    On Error Resume Next
    Load highscore(c)
    Load score(c)
    On Error GoTo 0
    
    highscore(c).Top = highscore(0).Top + c * highscore(c).Height
    highscore(c).Left = highscore(c).Left
    score(c).Left = score(c).Left
    score(c).Top = score(0).Top + c * score(c).Height
    
    highscore(c).Caption = ts.ReadLine
    score(c).Caption = ts.ReadLine
    
    highscore(c).Visible = True
    score(c).Visible = True
    c = c + 1
    If Not ts.AtEndOfStream And c < 11 Then GoTo more
    ts.Close
End Sub

Public Sub updatehighscore(newscore As Currency)
    'check to see if newscore is worth a high score
    'if so, then ask user for it's name, and add it too the right place in the file.
    scored = False
    Set ts = fso.OpenTextFile(fso.BuildPath(App.Path, "highscores.txt"), ForWriting, True)
    For a = 1 To highscore.UBound
        If score(a) < newscore And scored = False Then
            ts.WriteLine InputBox("You got a high score! Enter your name for the record books", "Lunar Lander", "Anonymous")
            ts.WriteLine newscore
            scored = True
        End If
        ts.WriteLine highscore(a)
        ts.WriteLine score(a)
    Next a
    ts.Close
    loadhighscores
End Sub

Private Sub Form_Queryunload(Cancel As Integer, UnloadMode As Integer)
    'clear all the sprites from memory when the user quits.
    For Each x In lib.Keys
        DeleteDC lib(x)
    Next x
    End
End Sub

