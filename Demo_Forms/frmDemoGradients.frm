VERSION 5.00
Begin VB.Form frmDemoGradients 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFEFEF&
   Caption         =   "   ___ Demo *Background Gradient Designs* ___"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
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
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   StartUpPosition =   2  'Bildschirmmitte
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   1245
      Top             =   435
      _ExtentX        =   847
      _ExtentY        =   847
      BackgroundGradient=   3
      BGWidth         =   210
      BGColorChange   =   30
      BGColor1        =   14737632
      BGColor2        =   16438204
      BGColor3        =   13237758
   End
   Begin VB.OptionButton optPercent 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Absolut  (Values < 0 ! )"
      Height          =   225
      Index           =   1
      Left            =   6300
      TabIndex        =   15
      Top             =   2775
      Width           =   2295
   End
   Begin VB.OptionButton optPercent 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Percent (Values 1-100)"
      Height          =   225
      Index           =   0
      Left            =   3780
      TabIndex        =   14
      Top             =   2775
      Value           =   -1  'True
      Width           =   2220
   End
   Begin VB.PictureBox picBGradType 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFEFEF&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   3765
      ScaleHeight     =   1380
      ScaleWidth      =   4440
      TabIndex        =   8
      Top             =   90
      Width           =   4470
      Begin VB.OptionButton optGradientType 
         BackColor       =   &H00EFEFEF&
         Caption         =   "0 - No Gradient"
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   3510
      End
      Begin VB.OptionButton optGradientType 
         BackColor       =   &H00EFEFEF&
         Caption         =   "1 - Two Color Gradient"
         Height          =   345
         Index           =   1
         Left            =   30
         TabIndex        =   11
         Top             =   375
         UseMaskColor    =   -1  'True
         Width           =   3510
      End
      Begin VB.OptionButton optGradientType 
         BackColor       =   &H00EFEFEF&
         Caption         =   "2 - Two Color Gradient plus Block"
         Height          =   345
         Index           =   2
         Left            =   30
         TabIndex        =   10
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   3510
      End
      Begin VB.OptionButton optGradientType 
         BackColor       =   &H00EFEFEF&
         Caption         =   "3 - Three Color Gradient"
         Height          =   345
         Index           =   3
         Left            =   30
         TabIndex        =   9
         Top             =   975
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   3510
      End
   End
   Begin VB.CommandButton btnRndColor 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Random Color 3 (Bottom)"
      Height          =   855
      Index           =   2
      Left            =   7125
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   3780
      Width           =   1110
   End
   Begin VB.CommandButton btnRndColor 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Random Color 2 (Middle)"
      Height          =   855
      Index           =   1
      Left            =   5445
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   3780
      Width           =   1110
   End
   Begin VB.CommandButton btnRndColor 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Random Color 1 (Top)"
      Height          =   855
      Index           =   0
      Left            =   3765
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   3780
      Width           =   1110
   End
   Begin VB.HScrollBar HScrollBorderColChange 
      Height          =   285
      LargeChange     =   20
      Left            =   3750
      Max             =   100
      Min             =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3315
      Value           =   30
      Width           =   4485
   End
   Begin VB.HScrollBar HScrollGradWidth 
      Height          =   315
      LargeChange     =   50
      Left            =   3750
      Max             =   800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Value           =   210
      Width           =   4485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please resize this form and play with parameters to check out the behavior."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   225
      TabIndex        =   17
      Top             =   3945
      Width           =   2445
   End
   Begin VB.Label lblBorderChangeType 
      BackStyle       =   0  'Transparent
      Caption         =   "Important for resizing the form:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3810
      TabIndex        =   16
      Top             =   2550
      Width           =   3045
   End
   Begin VB.Label lblNotice 
      BackStyle       =   0  'Transparent
      Caption         =   "And again, we are setting properties for demon- stration purposes only. No code is required to get the gradient functionality."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   195
      TabIndex        =   13
      Top             =   1560
      Width           =   2445
   End
   Begin VB.Label lblHint 
      BackStyle       =   0  'Transparent
      Caption         =   "The 'DrawTopDownGradient() function was the fastest compatible we found on a small competition on PSC/VB ;)"
      Height          =   690
      Left            =   3885
      TabIndex        =   7
      Top             =   4965
      Width           =   3555
   End
   Begin VB.Label lblColChange 
      BackStyle       =   0  'Transparent
      Caption         =   "Border Color Change 30%"
      Height          =   285
      Left            =   3825
      TabIndex        =   2
      Top             =   3060
      Width           =   3345
   End
   Begin VB.Label lblDescWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "Width 210 Pixels   (0 means: full width ! )"
      Height          =   285
      Left            =   3795
      TabIndex        =   0
      Top             =   1665
      Width           =   3570
   End
End
Attribute VB_Name = "frmDemoGradients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmDemoGradients.frm
'

Option Explicit
'
'
'



' ***************************************************************
' * All of this subs are just setting properties of WizzForm as *
' * you would do on design time. To play on your own just put   *
' * the WizzForm control onto an empty form, change some pro-   *
' * perties and start the application  to see the effects.      *
' ***************************************************************

Private Sub Form_Load()

    HScrollGradWidth.Max = Me.ScaleWidth
    HScrollGradWidth.Value = Me.ScaleWidth / 3
    
End Sub


Private Sub Form_Resize()
    
    ' Needs update on resizing of course
    HScrollGradWidth.Max = Me.ScaleWidth
    HScrollGradWidth.Value = Me.ScaleWidth / 3
    
End Sub

Private Sub optGradientType_Click(Index As Integer)
    
    ucWizzForm.BackgroundGradient = Index
    
End Sub


Private Sub HScrollGradWidth_Scroll()

    lblDescWidth.Caption = "Width " & HScrollGradWidth.Value & " Pixels   ( 0 means: Full width)"
    ucWizzForm.BG_Width = HScrollGradWidth.Value
        
End Sub


Private Sub optPercent_Click(Index As Integer)

    If Index = 0 Then
        lblColChange.Caption = "Border Color Change 30%"
        With HScrollBorderColChange
            .Min = 0
            .Max = 100
            .Value = 30
        End With
    Else
        lblColChange.Caption = "Border Color -100"
        With HScrollBorderColChange
            .Min = -1
            .Max = -500
            .Value = -100
        End With
    End If
    HScrollBorderColChange_Scroll
    
End Sub

Private Sub HScrollBorderColChange_Scroll()
    
    ucWizzForm.BG_ColorChange = HScrollBorderColChange.Value
    
    If optPercent(0).Value = True Then
        lblColChange.Caption = "Border Color Change " & HScrollBorderColChange.Value & "%"
    Else
        lblColChange.Caption = "Border Color Change " & HScrollBorderColChange.Value & " Pixels"
    End If
    
End Sub



Private Sub btnRndColor_Click(Index As Integer)

    If Index = 0 Then
        ucWizzForm.BGColor1 = RGB(256 * Rnd(1), 256 * Rnd(1), 256 * Rnd(1))
    
    ElseIf Index = 1 Then
        ucWizzForm.BGColor2 = RGB(256 * Rnd(1), 256 * Rnd(1), 256 * Rnd(1))
    
    Else
        ucWizzForm.BGColor3 = RGB(256 * Rnd(1), 256 * Rnd(1), 256 * Rnd(1))
    
    End If

End Sub


' #*#
