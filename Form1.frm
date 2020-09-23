VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   Caption         =   "SUPER PLAYER(by elie demien)"
   ClientHeight    =   4470
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   5520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "Form1.frx":1782
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6000
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   -30
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -890
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu rgr 
      Caption         =   "file"
      Begin VB.Menu sdfsdsdf 
         Caption         =   "open video"
      End
      Begin VB.Menu fwefwef 
         Caption         =   "open music"
      End
      Begin VB.Menu ewqeqw 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu wq 
      Caption         =   "options"
      Begin VB.Menu ewq 
         Caption         =   "mute on"
      End
      Begin VB.Menu fdgf 
         Caption         =   "volume control"
      End
      Begin VB.Menu ewfw 
         Caption         =   "play"
      End
      Begin VB.Menu ewe 
         Caption         =   "mute off"
      End
      Begin VB.Menu wqqwdqwd 
         Caption         =   "stop"
      End
      Begin VB.Menu fsdfsdsfsfsfs 
         Caption         =   "pause"
      End
   End
   Begin VB.Menu erwreww 
      Caption         =   "scan video"
      Begin VB.Menu weqwqweqweqeqe 
         Caption         =   "mpeg"
      End
      Begin VB.Menu ewqewq 
         Caption         =   "mpg"
      End
      Begin VB.Menu greeegeg 
         Caption         =   "avi"
      End
      Begin VB.Menu fwefwe 
         Caption         =   "mov"
      End
   End
   Begin VB.Menu erwr 
      Caption         =   "scan music"
      Begin VB.Menu fsdfsfsdf 
         Caption         =   "mp3"
      End
      Begin VB.Menu ewqeqe 
         Caption         =   "wav"
      End
      Begin VB.Menu asdadasd 
         Caption         =   "mid"
      End
   End
   Begin VB.Menu ewf 
      Caption         =   "search"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asdadasd_Click()
Call Load(Form8)
Form8.Show
End Sub

Private Sub ewe_Click()
MediaPlayer1.Mute = False
End Sub

Private Sub ewf_Click()
Call Load(Form10)
Form10.Show

End Sub

Private Sub ewfw_Click()
On Error Resume Next
MediaPlayer1.Play
End Sub

Private Sub ewq_Click()
MediaPlayer1.Mute = True
End Sub

Private Sub ewqeqe_Click()
Call Load(Form7)
Form7.Show
End Sub

Private Sub ewqeqw_Click()
Unload Me
End Sub

Private Sub ewqewq_Click()
Call Load(Form5)
Form5.Show
End Sub

Private Sub fdgf_Click()
On Error Resume Next
Shell "c:\windows\sndvol32.exe"
End Sub

Private Sub Form_Load()
MediaPlayer1.VideoBorder3D = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.MediaPlayer1.Height = Me.Height - 400
Me.MediaPlayer1.Width = Me.Width - 100
Me.Image1.Height = Me.Height - 1100
Me.Image1.Width = Me.MediaPlayer1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload Form9
Unload Form10

End Sub

Private Sub fsdfsdsfsfsfs_Click()
On Error Resume Next
MediaPlayer1.Pause
End Sub

Private Sub fsdfsfsdf_Click()
Call Load(Form6)
Form6.Show
End Sub

Private Sub fwefwe_Click()
Call Load(Form3)
Form3.Show
End Sub

Private Sub fwefwef_Click()
CommonDialog1.Filter = "mp3,wav,mid|*.mp3;*.wav;*.avi"
CommonDialog1.ShowOpen
MediaPlayer1.FileName = CommonDialog1.FileName
End Sub

Private Sub greeegeg_Click()
Call Load(Form4)
Form4.Show
End Sub

Private Sub sdfsdsd_Click()
Call Load(frmAbout)
frmAbout.Show
End Sub

Private Sub sdfsdsdf_Click()
CommonDialog1.Filter = "mpeg,mpg,mov,avi|*.mpeg;*.mov;*.avi;*.mpg"
CommonDialog1.ShowOpen
MediaPlayer1.FileName = CommonDialog1.FileName
End Sub

Private Sub weqwqweqweqeqe_Click()
Call Load(Form2)
Form2.Show
End Sub

Private Sub wqqwdqwd_Click()
On Error Resume Next
MediaPlayer1.Stop
End Sub
