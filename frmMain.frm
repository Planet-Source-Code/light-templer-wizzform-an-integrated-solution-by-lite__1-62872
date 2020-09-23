VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E6FCFD&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "  WizzForm Demo by Light Templer"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Background Gradients with Many Designs"
      Height          =   735
      Index           =   8
      Left            =   6135
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   1635
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Many Other Nice Little Things ..."
      Height          =   735
      Index           =   7
      Left            =   6135
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   3525
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Center Form in Workarea with Resizing"
      Height          =   735
      Index           =   6
      Left            =   4335
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   3525
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Save Position and Size to Reg or Ini"
      Height          =   735
      Index           =   5
      Left            =   2520
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   2580
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Resize Controls"
      Height          =   735
      Index           =   4
      Left            =   2520
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   3525
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Keep Aspect Ratio on Resizing"
      Height          =   735
      Index           =   3
      Left            =   6135
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   2580
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Additional Buttons in Caption Bar"
      Height          =   735
      Index           =   2
      Left            =   2520
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   1635
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Restricted Size"
      Height          =   735
      Index           =   1
      Left            =   4335
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   2580
      Width           =   1590
   End
   Begin VB.CommandButton btnShowDemoForm 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Many Additional Events"
      Height          =   735
      Index           =   0
      Left            =   4335
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   1635
      Width           =   1590
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   6795
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      CollapseButton  =   -1  'True
      StayOnTopButton =   -1  'True
      CollapseSmallSize=   29
      MsgHandle       =   54321
      BackgroundGradient=   3
      BGWidth         =   135
      BGColorChange   =   80
      BGColor2        =   16368031
      BGColor3        =   16761087
   End
   Begin VB.Line Line 
      Index           =   4
      X1              =   397
      X2              =   409
      Y1              =   196
      Y2              =   196
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   277
      X2              =   287
      Y1              =   233
      Y2              =   223
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   276
      X2              =   288
      Y1              =   133.333
      Y2              =   133.333
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   276
      X2              =   288
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   183
      X2              =   444
      Y1              =   44
      Y2              =   44
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "by Light Templer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4358
      TabIndex        =   15
      Top             =   855
      Width           =   1545
   End
   Begin VB.Label lblWizzForm 
      BackStyle       =   0  'Transparent
      Caption         =   "WizzForm 1.02"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   1
      Left            =   2775
      TabIndex        =   14
      Top             =   300
      Width           =   2355
   End
   Begin VB.Label lblWizzForm 
      BackStyle       =   0  'Transparent
      Caption         =   "WizzForm 1.02"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   405
      Index           =   0
      Left            =   2805
      TabIndex        =   13
      Top             =   285
      Width           =   2355
   End
   Begin VB.Shape Shape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   150
      Index           =   1
      Left            =   2520
      Top             =   1080
      Width           =   5220
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      Caption         =   "* An integrated      solution:  The       parts are work-    ing together."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   3
      Left            =   75
      TabIndex        =   11
      Top             =   3990
      Width           =   1635
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      Caption         =   "* All in one user-     control, no           additional bas      modul."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   2
      Left            =   75
      TabIndex        =   10
      Top             =   2820
      Width           =   1635
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      Caption         =   "* Subclassing in      IDE without the   crash. Thx to        Paul Caton!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   1
      Left            =   75
      TabIndex        =   9
      Top             =   1580
      Width           =   1650
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      Caption         =   "* No more code for     standard tasks.       Property setting      is enough."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   345
      Width           =   1830
   End
   Begin VB.Line Line1 
      X1              =   222
      X2              =   222
      Y1              =   223
      Y2              =   234
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmMain.frm
'

Option Explicit
'
'
'

' ==============================================================================
' = Here we just start the demo forms and handle a received window message the =
' = All informations & credits are at start of the usercontrols code.          =
' = Plz have a closer look there.                                              =
' ==============================================================================

Private Sub btnShowDemoForm_Click(Index As Integer)

    Select Case Index

        Case 0:     frmDemoEvents.Show vbModal

        Case 1:     frmDemoRestrictSize.Show vbModal
        
        Case 2:     frmDemoButtons.Show vbModal
        
        Case 3:     frmDemoAspectRatio.Show vbModal
        
        Case 4:     frmDemoResizeControlsPosition.Show vbModal
                
        Case 5:     frmDemoSavePosSize.Show vbModal
        
        Case 6:     frmDemoPositioning.Show vbModal
        
        Case 7:     frmDemoMisc.Show vbModal
        
        Case 8:     frmDemoGradients.Show vbModal
        
    End Select

End Sub


Private Sub ucWizzForm_ReceivedMessage(sMessage As String)
    
    MsgBox sMessage, vbInformation, " Just received this string from another app by WM_COPYDATA:"
    
End Sub

' #*#
