VERSION 5.00
Begin VB.Form frmDemoButtons 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "   ___ Demo *Additional Events* ___"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   6705
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   7785
      Left            =   975
      Picture         =   "frmDemoAdditionalButtons.frx":0000
      ScaleHeight     =   7785
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   1020
      Width           =   2895
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   405
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      CollapseButton  =   -1  'True
      StayOnTopButton =   -1  'True
      CollapseSmallSize=   83
      BackgroundGradient=   2
      BGColorChange   =   -60
      BGColor2        =   16368031
      BGColor3        =   16572365
   End
   Begin VB.Label lblHeadline 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Demo :   One or two additional buttons in forms title bar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   990
      TabIndex        =   2
      Top             =   195
      Width           =   2970
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stay on top"
      ForeColor       =   &H00808000&
      Height          =   270
      Left            =   4755
      TabIndex        =   1
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Collapse / expand"
      ForeColor       =   &H00800080&
      Height          =   390
      Left            =   5655
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmDemoButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'
'
'


' =================================================
' = All informations&credits are at start of the  =
' = usercontrols code. Plz have a look there.     =
' =================================================


Private Sub ucWizzForm_FormCollapse(flgShrink As Boolean)
    
    ' This event is raised before form changes its size.
    
End Sub

Private Sub ucWizzForm_FormStayOnTop(flgActiated As Boolean)

    ' This event is raised when additional button is pressed.

End Sub


' #*#
