VERSION 5.00
Begin VB.UserControl conttop 
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11385
   ScaleHeight     =   8490
   ScaleWidth      =   11385
   ToolboxBitmap   =   "cont.ctx":0000
   Begin VB.Image Image2 
      Height          =   6585
      Left            =   5280
      Picture         =   "cont.ctx":0312
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   6165
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2012 Kiran Pantha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Kiransoft Nepal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   3285
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build 3297"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KiranPlayer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3600
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "LicenseTo Kiransoft Software [Alpha Version]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   3375
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kiransoft's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1860
   End
   Begin VB.Image Image3 
      Height          =   2400
      Left            =   7320
      Picture         =   "cont.ctx":2F30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3960
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   5280
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   6330
      Left            =   0
      Picture         =   "cont.ctx":337A
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   11175
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2175
      Index           =   34
      Left            =   0
      Top             =   0
      Width           =   11415
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      Begin VB.Menu player 
         Caption         =   "About kiranPlayer"
      End
      Begin VB.Menu helpp 
         Caption         =   "Online Help"
      End
   End
End
Attribute VB_Name = "conttop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
