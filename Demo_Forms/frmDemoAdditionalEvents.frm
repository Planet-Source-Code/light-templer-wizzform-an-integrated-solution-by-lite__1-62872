VERSION 5.00
Begin VB.Form frmDemoEvents 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   Caption         =   "   ___ Demo *Additional Events* ___"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8325
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picBnewEvents 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   4980
      Picture         =   "frmDemoAdditionalEvents.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   3000
      TabIndex        =   4
      Tag             =   "|L-"
      Top             =   1440
      Width           =   3000
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   585
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00FDE4DB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Tag             =   "|RB"
      Top             =   1440
      Width           =   4230
   End
   Begin VB.Label lblHeadline 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo :   Additional events (e.g. resize the form or move it, switch to another application  ....)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1575
      TabIndex        =   5
      Top             =   240
      Width           =   5790
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Screenshot new events"
      ForeColor       =   &H00808000&
      Height          =   330
      Left            =   5040
      TabIndex        =   3
      Tag             =   "|L-"
      Top             =   1080
      Width           =   2940
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Event raised"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   315
      TabIndex        =   2
      Top             =   1140
      Width           =   2115
   End
   Begin VB.Label lblScreenShot 
      BackStyle       =   0  'Transparent
      Caption         =   "Screenshot new events"
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   5055
      TabIndex        =   1
      Tag             =   "|L-"
      Top             =   1065
      Width           =   2940
   End
End
Attribute VB_Name = "frmDemoEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmDemoAdditionalEvents.frm
'

Option Explicit
'
'
'

' =====================================================================
' = None of the subs here is for getting the functionality. They just =
' = demonstrate the possibilities. All informations&credits           =
' = are at start of the usercontrols code. Plz have a look there.     =
' =====================================================================


Private Sub AddToOutput(sMsg As String)

    With txtMsg
        .Text = .Text + sMsg + vbCrLf
        .SelStart = Len(.Text)
    End With

End Sub


Private Sub ucWizzForm_AppActivated()
    
    AddToOutput "Applikation activated"
    
End Sub

Private Sub ucWizzForm_AppDeactivated()
    
    AddToOutput "Applikation deactivated"
    
    MsgBox "You changed to a different application", vbInformation, " Event tracked:"
    
End Sub

Private Sub ucWizzForm_Error(lErrNo As Long, sErrMsg As String)

    AddToOutput " Error # " & lErrNo & "   " & sErrMsg

End Sub

Private Sub ucWizzForm_FormCollapse(flgShrink As Boolean)

    ' Used together with 'Additional buttons' in forms titlebar - look at other demo.

End Sub

Private Sub ucWizzForm_FormMoveSizeStart()

    AddToOutput "Start: Form moving or sizing"

End Sub

Private Sub ucWizzForm_FormMoveSizeEnd()

    AddToOutput "End: Form moving or sizing"

End Sub

Private Sub ucWizzForm_FormMoving()

    AddToOutput "Form moving"

End Sub


Private Sub ucWizzForm_FormResizing(Edge As enWMSZ)
    
    AddToOutput "Form sizing - Edge = " & Edge
    
End Sub

Private Sub ucWizzForm_FormStayOnTop(flgActiated As Boolean)
    
    ' Used together with 'Additional buttons' in forms titlebar - look at other demo.
    
End Sub

Private Sub ucWizzForm_MouseEntersForm()
    
    AddToOutput "Mouse over form"
    
End Sub

Private Sub ucWizzForm_MouseLeavesForm()
    
    AddToOutput "Mouse leaves form"
    
End Sub

Private Sub ucWizzForm_ScreenResChanged(NewWidth As Long, NewHeight As Long, NewColorDepth As Long)

    AddToOutput "New screen resolution:  " & _
            NewWidth & " x " & NewHeight & _
            " / " & _
            NewColorDepth & " Bits/Pixel"

End Sub

Private Sub ucWizzForm_SystemColorChange(flgEffectOfThemeChange As Boolean)

    AddToOutput IIf(flgEffectOfThemeChange = True, _
            "Theme & system colors change", _
            "System colors change")

End Sub

Private Sub ucWizzForm_ThemeChange()
    
    AddToOutput "Theme change"
    
End Sub

' #*#
