VERSION 5.00
Begin VB.UserControl playcontrol 
   BackColor       =   &H0000FF00&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   915
   ScaleHeight     =   495
   ScaleWidth      =   915
   Begin VB.Line Line3 
      X1              =   480
      X2              =   720
      Y1              =   360
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   600
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   480
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   495
      Index           =   1
      Left            =   360
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   495
      Index           =   0
      Left            =   240
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "playcontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
