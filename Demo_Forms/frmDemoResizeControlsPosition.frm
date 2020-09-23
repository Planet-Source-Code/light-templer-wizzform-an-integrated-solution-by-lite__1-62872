VERSION 5.00
Begin VB.Form frmDemoResizeControlsPosition 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   Caption         =   "   ___ Demo *Resize Controls* ___"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
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
   ScaleHeight     =   5985
   ScaleWidth      =   8865
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox pivBTag 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3060
      Left            =   315
      Picture         =   "frmDemoResizeControlsPosition.frx":0000
      ScaleHeight     =   3060
      ScaleWidth      =   4065
      TabIndex        =   22
      Top             =   1020
      Width           =   4065
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3570
      Index           =   2
      Left            =   4710
      ScaleHeight     =   3540
      ScaleWidth      =   3765
      TabIndex        =   7
      Tag             =   "|R-"
      Top             =   1020
      Width           =   3795
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "|LB"
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   12
         Left            =   2895
         TabIndex        =   20
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "MyValues"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   11
         Left            =   2145
         TabIndex        =   19
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* A - at one of the two postions means:           Hands off - Nothing to change."
         Height          =   465
         Index           =   10
         Left            =   90
         TabIndex        =   18
         Top             =   3075
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* B  - Move BOTTOM border (change size)"
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   17
         Top             =   2835
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* C  - Move TOP border (center control)"
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   16
         Top             =   2610
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* T  - Move TOP border (move control)"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   15
         Top             =   2415
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* R  - Move RIGHT border (change size)"
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   14
         Top             =   2130
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* M  - Move LEFT border (center ctrl)"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   1920
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "After the | there must be 2 more chars:"
         Height          =   240
         Index           =   4
         Left            =   210
         TabIndex        =   12
         Top             =   1485
         Width           =   3330
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "* L  - Move LEFT border (move control)"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   1710
         Width           =   3480
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   10
         Top             =   810
         Width           =   135
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "A | delimits other values stored in tag from the resizing infos. An example for a tag value could be    MyValues|LB"
         Height          =   645
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   3270
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "The TAG property of a control is used to hold the information which way resizing is to do."
         Height          =   645
         Index           =   0
         Left            =   255
         TabIndex        =   8
         Top             =   30
         Width           =   3330
      End
   End
   Begin VB.PictureBox Picture 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   1
      Left            =   300
      ScaleHeight     =   1170
      ScaleWidth      =   8145
      TabIndex        =   0
      Tag             =   "|BR"
      Top             =   4695
      Width           =   8205
      Begin VB.CommandButton Command 
         BackColor       =   &H00CE8359&
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   6030
         Style           =   1  'Grafisch
         TabIndex        =   3
         Tag             =   "|LC"
         Top             =   225
         Width           =   1530
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00CE8359&
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3285
         Style           =   1  'Grafisch
         TabIndex        =   2
         Tag             =   "|MC"
         Top             =   225
         Width           =   1530
      End
      Begin VB.CommandButton Command 
         BackColor       =   &H00CE8359&
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   555
         Style           =   1  'Grafisch
         TabIndex        =   1
         Tag             =   "|-C"
         Top             =   225
         Width           =   1530
      End
      Begin VB.Label lblTag 
         Alignment       =   2  'Zentriert
         BackColor       =   &H006BD2FE&
         Caption         =   "Tag:  |LC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   6030
         TabIndex        =   6
         Tag             =   "|LC"
         Top             =   735
         Width           =   1515
      End
      Begin VB.Label lblTag 
         Alignment       =   2  'Zentriert
         BackColor       =   &H006BD2FE&
         Caption         =   "Tag:  |MC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3285
         TabIndex        =   5
         Tag             =   "|MC"
         Top             =   735
         Width           =   1515
      End
      Begin VB.Label lblTag 
         Alignment       =   2  'Zentriert
         BackColor       =   &H006BD2FE&
         Caption         =   "Tag:  |-C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   555
         TabIndex        =   4
         Tag             =   "|-C"
         Top             =   735
         Width           =   1515
      End
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   315
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblHeadline 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo :   Auto resize controls when form size is changing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1065
      TabIndex        =   21
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmDemoResizeControlsPosition"
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


' No line of code needed ;)


' #*#
