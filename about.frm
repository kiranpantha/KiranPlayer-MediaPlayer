VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H0000FF00&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8490
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "about.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "about.frx":000C
   ScaleHeight     =   8490
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image4 
      Height          =   1560
      Left            =   240
      Picture         =   "about.frx":0316
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2280
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Website:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   10560
      MouseIcon       =   "about.frx":0620
      MousePointer    =   99  'Custom
      Picture         =   "about.frx":092A
      Top             =   120
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Studing at Class XII in New Horizon College"
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
      TabIndex        =   10
      Top             =   4800
      Width           =   6450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " www.kiranpantha.com.np"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   5760
      Width           =   3885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: info@kiranpantha.com.np"
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
      TabIndex        =   8
      Top             =   5280
      Width           =   5025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address : Shankarnagar - 5 rupandehi Nepal"
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
      TabIndex        =   7
      Top             =   3720
      Width           =   6705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name: Kiran Pantha"
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
      TabIndex        =   6
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   6345
      Left            =   5280
      Picture         =   "about.frx":0D47
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   6195
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   5280
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   2400
      Left            =   7320
      Picture         =   "about.frx":3965
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3960
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
      TabIndex        =   5
      Top             =   0
      Width           =   1860
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
      TabIndex        =   4
      Top             =   8040
      Width           =   3375
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
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   3600
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   3285
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2009-2013 Kiran Pantha"
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
      TabIndex        =   0
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   8535
      Index           =   34
      Left            =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

lblCopyright = "Copyright 2009-" & Year(Date$) & " Kiran Pantha"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
MsgBox "Error in opening Mail program or Browser" & Chr(10) & "Error 0x158918", vbCritical, "Error 0x158918"
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Shell "C:\Program Files\Internet Explorer\iexplore.exe mailto:info@kiranpantha.com.np", vbNormalFocus

End Sub

Private Sub Label4_Click()
Shell "C:\Program Files\Internet Explorer\iexplore.exe http://kiranpantha.com.np", vbNormalFocus
End Sub
