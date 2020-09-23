VERSION 5.00
Begin VB.Form frmDemoAspectRatio 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   Caption         =   "   ___ Demo *Keep aspect ratio on resizing* ___"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picBTop 
      Appearance      =   0  '2D
      BackColor       =   &H00CE8359&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   1005
      ScaleHeight     =   840
      ScaleWidth      =   5040
      TabIndex        =   5
      Tag             =   "|M-"
      Top             =   60
      Width           =   5070
      Begin VB.Label lblHeadline 
         Appearance      =   0  '2D
         BackColor       =   &H00CE8359&
         Caption         =   $"frmDemoAspectRatio.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   210
         TabIndex        =   6
         Top             =   60
         Width           =   4770
      End
   End
   Begin VB.PictureBox pivBRight 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   6480
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   3
      Tag             =   "|LB"
      Top             =   990
      Width           =   405
      Begin VB.Label lblHeightNumber 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFC0FF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   4
         Tag             =   "|-C"
         Top             =   1320
         Width           =   405
      End
   End
   Begin VB.PictureBox picBKeepAspect 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2970
      Left            =   1530
      Picture         =   "frmDemoAspectRatio.frx":0092
      ScaleHeight     =   2940
      ScaleWidth      =   4020
      TabIndex        =   0
      Tag             =   "|MC"
      Top             =   1215
      Width           =   4050
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   315
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      KeepAspectRatio =   -1  'True
      BackgroundGradient=   1
      BGColorChange   =   -65
      BGColor1        =   16368031
   End
   Begin VB.PictureBox picBBottom 
      Appearance      =   0  '2D
      BackColor       =   &H00E6FCFD&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   459
      TabIndex        =   1
      Tag             =   "|TR"
      Top             =   4635
      Width           =   6885
      Begin VB.Label lblWidthNumber 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFC0FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         TabIndex        =   2
         Tag             =   "|M-"
         Top             =   0
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmDemoAspectRatio"
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


' No code neccessary ;)

' Reposition of controls:   Plz look at their 'Tag' property
'                           Closer description in text of demo form.



' #*#


