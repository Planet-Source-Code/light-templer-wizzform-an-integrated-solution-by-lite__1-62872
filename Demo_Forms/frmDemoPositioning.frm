VERSION 5.00
Begin VB.Form frmDemoPositioning 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   Caption         =   "   ___ Demo *Positioning* ___"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnPos4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gap to border: 200 pixels"
      Height          =   645
      Left            =   3435
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   2550
      Width           =   1440
   End
   Begin VB.CommandButton btnPos3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gap to border: 30 pixels"
      Height          =   645
      Left            =   1080
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   2550
      Width           =   1440
   End
   Begin VB.CommandButton btnPos2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gap to border: 0"
      Height          =   645
      Left            =   3435
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   1215
      Width           =   1440
   End
   Begin VB.CommandButton btnPos1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Center - No resizing"
      Height          =   645
      Left            =   1080
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   1200
      Width           =   1440
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   330
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblHeadline 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo :   Resize form and put it centered into useable area of the screen with one function call"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   165
      Width           =   4890
   End
End
Attribute VB_Name = "frmDemoPositioning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'
'   frmDemoPositioning.frm
'

' *****************************************************************
' * Examples of calling one of the usefull subs to show centering *
' * with and without resizing of the form.                        *
' *****************************************************************
'
'
'

Private Sub btnPos1_Click()
    '  -1:  Just center form - no resizing.
    '       (btw: -1 is default (optional) parameter)
    
    ucWizzForm.CenterFormInWorkArea -1
    
End Sub

Private Sub btnPos2_Click()
    '   0:  Resize form to fully fill the free desktop area.
    '       Taskbar and correct registered additional bars like
    '       MS Office bar will not covered.
    
    ucWizzForm.CenterFormInWorkArea 0
    
End Sub

Private Sub btnPos3_Click()
    '  30:  Here we center the form considering desktop bars and we
    '       keep a small distance to all borders 30 pixels wide.
    
    ucWizzForm.CenterFormInWorkArea 30
    
End Sub

Private Sub btnPos4_Click()
    ' 200:  Here we center the form considering desktop bars and we
    '       keep a large distance to all borders 200pixels wide.
    
    ucWizzForm.CenterFormInWorkArea 200
    
End Sub

' #*#
