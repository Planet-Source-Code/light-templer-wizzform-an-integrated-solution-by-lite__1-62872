VERSION 5.00
Begin VB.Form frmDemoRestrictSize 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   Caption         =   "   ___ Demo *Restricted Size* ___"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   5955
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4440
      Index           =   1
      Left            =   1095
      Picture         =   "frmDemoRestrictSize.frx":0000
      ScaleHeight     =   4440
      ScaleWidth      =   3180
      TabIndex        =   0
      Top             =   1185
      Width           =   3180
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   390
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      FormMinWidth    =   400
      FormMinHeight   =   400
      FormMinWidth    =   400
      FormMinHeight   =   400
      FormMaxWidth    =   600
      FormMaxHeight   =   500
   End
   Begin VB.Label lblHeadline 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Demo :   Restrict forms size to a min / max size"
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
      Left            =   1050
      TabIndex        =   1
      Top             =   345
      Width           =   4710
   End
End
Attribute VB_Name = "frmDemoRestrictSize"
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

' Reposition of controls:  Look at their 'Tag' property
'                          and open 'Resize control' demo.


' #*#
