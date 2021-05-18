VERSION 5.00
Begin VB.Form start 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "startatfirst.frx":0000
   ScaleHeight     =   3975
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ope 
      Interval        =   10
      Left            =   4320
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1440
   End
   Begin VB.CommandButton Top 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   3915
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton Top 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   3915
      Index           =   12
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton Top 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Top 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   1080
      ScaleHeight     =   2415
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   720
      Width           =   2775
      Begin VB.Image Image2 
         Height          =   840
         Left            =   0
         Picture         =   "startatfirst.frx":A8C42
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
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
         Height          =   1575
         Left            =   0
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label kiranpl 
         BackStyle       =   0  'Transparent
         Caption         =   "KiranPlayer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i
Private Sub Form_Load()
i = 1
End Sub

Private Sub ope_Timer()
k$ = "KiranPlayer is a simple media Player by Kiran pantha.Which is capable of playing different media files"
lbl = Left$(k$, i)
i = i + 1
If i >= Len(k$) Then Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
player.Show
Unload Me
End Sub


