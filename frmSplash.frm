VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Media File"
   ClientHeight    =   5535
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   10125
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close it"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FF00&
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9840
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmSplash.frx":030A
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         Picture         =   "frmSplash.frx":367C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6480
         Top             =   360
      End
      Begin VB.FileListBox files 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3690
         Left            =   3480
         TabIndex        =   3
         Top             =   840
         Width           =   6135
      End
      Begin VB.DirListBox dire 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3690
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.DriveListBox driv 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   525
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
If LCase(Right$(files.FileName, 3)) = "mp3" Or LCase(Right$(files.FileName, 3)) = "mp4" Or LCase(Right$(files.FileName, 3)) = "wav" Or LCase(Right$(files.FileName, 3)) = "avi" Or LCase(Right$(files.FileName, 3)) = "wma" Or LCase(Right$(files.FileName, 3)) = "mpg" Then
player.gk.URL = dire.Path & "/" & files.FileName
If LCase(Right$(files.FileName, 3)) = "mp3" Or LCase(Right$(files.FileName, 3)) = "wav" Or LCase(Right$(files.FileName, 3)) = "wma" Then
player.gk.uiMode = "invisible"
Else
player.gk.uiMode = "none"
player.Height = 8445
player.Width = 8145
End If
player.Caption = "KiranPlayer " & App.Major & "." & App.Minor & "." & App.Revision & " - " & player.gk.currentMedia.Name
Me.Hide
Else
MsgBox "Error in media file" & Chr(10) & "The given file is not a Mp3 or mp4 or Wav or avi or wmv or mpg format", vbCritical + vbOKOnly, "Error"
End If

End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub Timer1_Timer()
On Error GoTo endderr
dire.Path = driv
files.Path = dire
endderr:
End Sub
