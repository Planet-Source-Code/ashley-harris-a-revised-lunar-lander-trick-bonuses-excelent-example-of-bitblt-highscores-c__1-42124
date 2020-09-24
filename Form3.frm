VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf 
      Height          =   4470
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5970
      _cx             =   10530
      _cy             =   7885
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Opaque"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    swf.Movie = fso.BuildPath(App.Path, "intro.swf")
    swf.Play
End Sub

Private Sub swf_FSCommand(ByVal command As String, ByVal args As String)
    swf.Stop
    Form2.Show
    Me.Hide
    
End Sub
