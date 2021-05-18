VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form player 
   BackColor       =   &H00FFFFFF&
   Caption         =   "KiranPlayer"
   ClientHeight    =   7650
   ClientLeft      =   3960
   ClientTop       =   1845
   ClientWidth     =   8010
   Icon            =   "mainbody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   8010
   Begin VB.Frame menukey 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   8055
      Begin MSComDlg.CommonDialog Openit 
         Left            =   7080
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open The media file to play"
         InitDir         =   "/"
      End
      Begin VB.CommandButton play 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         Picture         =   "mainbody.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Play"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton pause 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         Picture         =   "mainbody.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Pause"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton ff 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2880
         Picture         =   "mainbody.frx":0B64
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fast Forward"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton stopp 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2280
         Picture         =   "mainbody.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Stop"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton fr 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1680
         Picture         =   "mainbody.frx":11C2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fast Rewind"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton fs 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3480
         Picture         =   "mainbody.frx":14F6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Full Screen"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton info 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4680
         Picture         =   "mainbody.frx":19F3
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Media Information"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton mute 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4080
         Picture         =   "mainbody.frx":1E77
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Mute"
         Top             =   1440
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   7560
         Top             =   720
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   7560
         Top             =   360
      End
      Begin VB.Timer ope 
         Interval        =   100
         Left            =   7560
         Top             =   1080
      End
      Begin MSComctlLib.Slider playslide 
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   4
         TickFrequency   =   0
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         ToolTipText     =   "Volume"
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   100
         SelStart        =   50
         TickStyle       =   3
         Value           =   50
      End
      Begin VB.CommandButton statu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   8
         Top             =   720
         Width           =   7335
      End
      Begin VB.Label command1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Open Media File"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   240
         Picture         =   "mainbody.frx":233F
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   5100
      End
      Begin VB.Label vol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label stat 
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
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.CommandButton mutebef 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6360
      MaskColor       =   &H80000001&
      Picture         =   "mainbody.frx":3591
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton muteaft 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6840
      MaskColor       =   &H80000001&
      Picture         =   "mainbody.frx":3A59
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "mainbody.frx":3FA8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   555
   End
   Begin VB.Image bckimg 
      Height          =   4755
      Left            =   0
      Picture         =   "mainbody.frx":42B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8040
   End
   Begin WMPLibCtl.WindowsMediaPlayer gk 
      Height          =   4680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14076
      _cy             =   8255
   End
   Begin VB.Menu kiran 
      Caption         =   "kiransoft"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu opennew 
         Caption         =   "Open Media File"
         Begin VB.Menu viah 
            Caption         =   "via Harddisk or removeable Media"
            Begin VB.Menu kpdia 
               Caption         =   "Via KiranPlayer's Dialog"
            End
            Begin VB.Menu windia 
               Caption         =   "Via Windows Dialog"
            End
         End
         Begin VB.Menu viai 
            Caption         =   "via internet or Local Newtwork"
         End
      End
      Begin VB.Menu pl 
         Caption         =   "Play"
      End
      Begin VB.Menu ause 
         Caption         =   "Pause"
      End
      Begin VB.Menu top 
         Caption         =   "Stop"
      End
      Begin VB.Menu fastf 
         Caption         =   "Fast Forward"
      End
      Begin VB.Menu fastr 
         Caption         =   "Fast Rewind"
      End
      Begin VB.Menu fulscr 
         Caption         =   "Full screen"
      End
      Begin VB.Menu volumm 
         Caption         =   "Volume"
         Begin VB.Menu vvp 
            Caption         =   "Volume -"
         End
         Begin VB.Menu vv 
            Caption         =   "Volume +"
         End
      End
      Begin VB.Menu abt 
         Caption         =   "About KiranPlayer"
      End
   End
   Begin VB.Menu fil 
      Caption         =   "File"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu kir 
         Caption         =   "About KiranPlayer"
         Shortcut        =   ^A
      End
      Begin VB.Menu onhelp 
         Caption         =   "Online help"
      End
   End
End
Attribute VB_Name = "player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filevol
Private Sub abt_Click()
about.Show
End Sub

Private Sub ause_Click()
Call pause_Click
End Sub

Private Sub bckimg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub Command1_Click()
Openit.ShowOpen

End Sub


Private Sub fastr_Click()
Call fr_Click
End Sub

Private Sub ff_Click()
gk.Controls.fastForward
End Sub

Private Sub fil_Click()
PopupMenu kiran
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
MsgBox Source
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
MsgBox CmdStr
End Sub

Private Sub Form_Load()
On Error GoTo ennd
Rem start
gk.Height = Me.Height - menukey.Height
gk.Width = Me.Width - 100
bckimg.Height = gk.Height
bckimg.Width = gk.Width
menukey.Width = Me.Width
playslide.Width = Me.Width
menukey.Top = gk.Height
Rem end
Dim gk_url$, playslide_Max, playslide_Value
Caption = "KiranPlayer " & App.Major & "." & App.Minor & "." & App.Revision
If Dir("playstate.kp") <> "" Then
Open "playstate.kp" For Input As #2
Input #2, playslide_Max, playslide_Value, gk_url$
Close #2
If playslide_Max <> 0 And playslide_Value <> 0 And gk_url$ <> "" Then
If MsgBox("The last media file has not ended " & gk_url$ & Chr(10) & "Want to continue", vbYesNo, "Continue...") = vbYes Then
gk.URL = gk_url$
If LCase(Right$(gk_url$, 3)) = "mp3" Or LCase(Right$(gk_url$, 3)) = "mp4" Or LCase(Right$(gk_url$, 3)) = "wav" Or LCase(Right$(gk_url$, 3)) = "avi" Or LCase(Right$(gk_url$, 3)) = "wma" Or LCase(Right$(gk_url$, 3)) = "mpg" Then
If LCase(Right$(gk_url$, 3)) = "mp3" Or LCase(Right$(gk_url$, 3)) = "wav" Or LCase(Right$(gk_url$, 3)) = "wma" Then
player.gk.uiMode = "invisible"
Else
player.gk.uiMode = "none"
End If
player.Caption = "KiranPlayer " & App.Major & "." & App.Minor & "." & App.Revision & " - " & gk_url$
Me.Hide
Else
MsgBox "Error in media file" & Chr(10) & "The given file is not a Mp3 or mp4 or Wav or avi or wmv or mpg format", vbCritical + vbOKOnly, "Error"
End If
playslide.Max = playslide_Max
gk.Controls.currentPosition = playslide_Value
End If
End If
End If
If Dir("Volume.kp") <> "" Then
Open "volume.kp" For Input As #1
Input #1, valueofvol
Close #1
 Slider1.Value = valueofvol
 End If
ennd:
End Sub



Private Sub Form_Resize()
On Error GoTo endd
gk.Height = Me.Height - menukey.Height
gk.Width = Me.Width - 100
bckimg.Height = gk.Height
bckimg.Width = gk.Width
menukey.Width = Me.Width
playslide.Width = Me.Width
menukey.Top = gk.Height
endd:
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub fr_Click()
gk.Controls.fastReverse
End Sub

Private Sub fs_Click()
On Error GoTo last
MsgBox "Double Click The video Area for full screen" & Chr(10) & "To exit click escape[esc] button", vbInformation, "Info"
Exit Sub
last:
MsgBox "Error 0x0003d65 has occured ." & Chr(10) & "There was problem in starting fullscreen", vbCritical, "Error in starting fullscreen"
End Sub

Private Sub fulscr_Click()
Call fs_Click
End Sub

Private Sub GK_MouseDown(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
If nButton = 2 Then PopupMenu kiran
End Sub

Private Sub info_Click()
On Error GoTo enddd
If gk.URL <> "" Then
If gk.settings.defaultFrame = "" Then frameRate = "Unknown"
If gk.currentMedia.imageSourceHeight <> 0 Then
demension = "Demension : " & gk.currentMedia.imageSourceHeight & "X" & gk.currentMedia.imageSourceWidth
Else
demension = "Demension : Unknown"
End If
MsgBox "Media Information" & Chr(10) & Chr(10) & "Name of Media : " & gk.currentMedia.Name & Chr(10) & Chr(10) & "Size of media : " & Round(FileLen(gk.URL) / 1024, 2) & " KB" & Chr(10) & Chr(10) & Chr(10) & Chr(10) & "Bitrate : " & Round(gk.network.bitRate / 1000, 2) & " KBPS" & Chr(10) & Chr(10) & "Location of file: " & gk.currentMedia.sourceURL & Chr(10) & Chr(10) & "Frame Rate : " & frameRate & Chr(10) & Chr(10) & demension & Chr(10) & Chr(10) & "File type: " & Right$(gk.URL, 3) & " file format", vbInformation + vbOKOnly, "Media Information"
Else
MsgBox "Please Open a media file", vbInformation + vbOKOnly
End If
Exit Sub
enddd:
If LCase(Left$(gk.URL, 4)) = "http" Then MsgBox "File Information Access Denied by Webserver", vbInformation, "Sorry! Access Denined": Exit Sub
MsgBox "It seems that the device containg media file is ejected from the computer. Insert to play it again", vbCritical + vbOKOnly, "Ejected"
End Sub

Private Sub info_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub kir_Click()
about.Show
End Sub

Private Sub kpdia_Click()
frmSplash.Show
End Sub

Private Sub menukey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub mute_Click()
If gk.settings.mute = True Then
If Dir(App.Path & "\muted.kp") <> "" Then
Open App.Path & "\muted.kp" For Input As #158
Input #158, volumee
Close #158
Slider1.Value = volumee
If Slider1.Value = 0 Then Slider1.Value = 50
Else
Slider1.Value = 50
End If
mute.Picture = mutebef.Picture
Slider1.Enabled = True
gk.settings.mute = False
Else
mute.Picture = muteaft.Picture
Open App.Path & "\muted.kp" For Output As #159
Write #159, Slider1.Value
Close #159
Slider1.Value = 0
Slider1.Enabled = False
gk.settings.mute = True
End If
End Sub

Private Sub onhelp_Click()
Shell "C:\Program Files\Internet Explorer\iexplore.exe http://kiranpantha.com.np/help.php?val=player&ver=" & App.Major & "." & App.Minor, vbNormalFocus
End Sub

Private Sub ope_Timer()

statu.Caption = Left(gk.Status, 60) & "...."
statu.ToolTipText = gk.Status
If Openit.FileName <> "" Then
kiranopen$ = Openit.FileName
Openit.FileName = ""
End If
If kiranopen$ <> "" Then
If LCase(Right$(kiranopen$, 3)) = "mp3" Or LCase(Right$(kiranopen$, 3)) = "mp4" Or LCase(Right$(kiranopen$, 3)) = "wav" Or LCase(Right$(kiranopen$, 3)) = "avi" Or LCase(Right$(kiranopen$, 3)) = "wma" Or LCase(Right$(kiranopen$, 3)) = "mpg" Then
player.gk.URL = kiranopen$
If LCase(Right$(kiranopen$, 3)) = "mp3" Or LCase(Right$(kiranopen$, 3)) = "wav" Or LCase(Right$(kiranopen$, 3)) = "wma" Then
player.gk.uiMode = "invisible"
Else
player.gk.uiMode = "none"
player.Height = 8445
player.Width = 8145
End If
player.Caption = "KiranPlayer " & App.Major & "." & App.Minor & "." & App.Revision & " - " & player.gk.currentMedia.Name
Else
MsgBox "Error in media file" & Chr(10) & "The given file is not a Mp3 or mp4 or Wav or avi or wmv or mpg format", vbCritical + vbOKOnly, "Error"
End If
End If
End Sub

Private Sub pause_Click()
gk.Controls.pause
End Sub

Private Sub pl_Click()
Call play_Click
End Sub

Private Sub play_Click()
gk.Controls.play
End Sub

Private Sub play_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub playcontrol1_GotFocus()

End Sub

Private Sub playslide_Change()
Open App.Path & "\playstate.kp" For Output As #2
Write #2, playslide.Max, playslide.Value, gk.URL
Close #2
End Sub

Private Sub playslide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = False

End Sub

Private Sub playslide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
gk.Controls.currentPosition = playslide.Value
Timer2.Enabled = True
End Sub



Private Sub Slider1_Change()
Open App.Path & "\volume.kp" For Output As #1
Write #1, Slider1.Value
Close #1
End Sub

Private Sub stat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub statu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub stopp_Click()
gk.Controls.stop
End Sub

Private Sub Timer1_Timer()
On Error GoTo enndd
vol.Left = Me.Width - vol.Width - 200
Slider1.Left = Me.Width - Slider1.Width - 200
If menukey.Top < 1000 And WindowState = 0 Then Me.Height = menukey.Height
gk.settings.volume = Slider1.Value
vol.Caption = "Volume : " & gk.settings.volume
enndd:
End Sub

Private Sub Timer2_Timer()
On Error GoTo ennddd
If Me.Width < 8250 Then Me.Width = 8250
If gk.uiMode = "invisible" And WindowState = 2 Then WindowState = 0
If gk.URL <> "" Then
If gk.uiMode = "invisible" And WindowState <> 2 Then
menukey.Top = 0
End If
mute.Visible = True
bckimg.Visible = False
play.Visible = True
pause.Visible = True
ff.Visible = True
stopp.Visible = True
fr.Visible = True
fs.Visible = True
info.Visible = True
statu.Visible = True
stat.Caption = "Total time of Media:" & gk.currentMedia.durationString & " - Time taken " & gk.Controls.currentPositionString
If gk.currentMedia.duration <> 0 Then playslide.Max = Int(gk.currentMedia.duration) - 1
playslide.Value = Int(gk.Controls.currentPosition)
Else
star1.Visible = False
star2.Visible = False
star3.Visible = False
star4.Visible = False
star4.Visible = False
mute.Visible = False
play.Visible = False
pause.Visible = False
ff.Visible = False
stopp.Visible = False
fr.Visible = False
fs.Visible = False
info.Visible = False
statu.Visible = False
bckimg.Visible = True
End If
ennddd:
End Sub



Private Sub top_Click()
Call stopp_Click
End Sub


Private Sub viai_Click()
On Error GoTo ennda
filenam$ = InputBox("Enter the long URL of the site from which you want to listen Music" & Chr(10) & "Example : http://kiranpantha.tk/songs/faraway.mp4", "Open through Network")
If LCase(Right$(filenam$, 3)) = "mp3" Or LCase(Right$(filenam$, 3)) = "mp4" Or LCase(Right$(filenam$, 3)) = "wav" Or LCase(Right$(filenam$, 3)) = "avi" Or LCase(Right$(filenam$, 3)) = "wma" Or LCase(Right$(filenam$, 3)) = "mpg" Then
player.gk.URL = filenam$
If LCase(Right$(filenam$, 3)) = "mp3" Or LCase(Right$(filenam$, 3)) = "wav" Then
player.gk.uiMode = "invisible"
Else
player.gk.uiMode = "none"
End If
Else
MsgBox "Error in media file" & Chr(10) & "The given file is not a Mp3 or mp4 or Wav or avi or wmv or mpg format or no Information was written", vbCritical + vbOKOnly, "Error"
End If
ennda:
End Sub

Private Sub vol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu kiran
End Sub

Private Sub vv_Click()
Slider1.Value = Slider1.Value + 10
End Sub

Private Sub vvp_Click()
Slider1.Value = Slider1.Value - 10
End Sub

Private Sub windia_Click()
Openit.ShowOpen
End Sub
