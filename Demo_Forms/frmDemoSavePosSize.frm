VERSION 5.00
Begin VB.Form frmDemoSavePosSize 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   Caption         =   "   ___ Demo *restore size and position* ___"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7575
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   180
      Picture         =   "frmDemoSavePosSize.frx":0000
      ScaleHeight     =   3330
      ScaleWidth      =   4065
      TabIndex        =   3
      Top             =   1095
      Width           =   4065
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3300
      Index           =   2
      Left            =   4545
      ScaleHeight     =   3270
      ScaleWidth      =   2760
      TabIndex        =   0
      Tag             =   "|RB"
      Top             =   1110
      Width           =   2790
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDemoSavePosSize.frx":2C0B4
         Height          =   1815
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   135
         Width           =   2535
      End
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   195
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      SavePosition    =   -1  'True
      SaveSize        =   -1  'True
   End
   Begin VB.Label lblHeadline 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemoSavePosSize.frx":2C1AC
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   945
      TabIndex        =   2
      Top             =   120
      Width           =   5910
   End
End
Attribute VB_Name = "frmDemoSavePosSize"
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


' No code neccessary ;) - Not even for the registry
'                         or ini file handling.


' #*#

