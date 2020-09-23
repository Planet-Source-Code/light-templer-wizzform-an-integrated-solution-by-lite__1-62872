VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   " Demo Sending a simple string message from app to app (form to form)"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtCaptionOfDestWindow 
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Text            =   "  WizzForm Demo by Light Templer"
      Top             =   195
      Width           =   6195
   End
   Begin VB.CommandButton btnSendMsg 
      Caption         =   "Send message"
      Default         =   -1  'True
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   1515
      Width           =   6195
   End
   Begin VB.TextBox txtMsg 
      Height          =   330
      Left            =   240
      MaxLength       =   254
      TabIndex        =   0
      Text            =   "Your demo message here!"
      Top             =   720
      Width           =   6195
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'
'   frmMain
'

Option Explicit


Private Type tpCOPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type

Private Declare Function API_FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long

Private Declare Function API_SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long

Private Declare Sub API_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef lPtrDest As Any, _
         ByRef lPtrSrc As Any, _
         ByVal lLength As Long)

Private Const WM_COPYDATA = &H4A
'
'
'

Private Sub btnSendMsg_Click()
    
    Dim cds                 As tpCOPYDATASTRUCT
    Dim hWnd                As Long
    Dim arrBytBuf(1 To 255) As Byte


    hWnd = API_FindWindow(vbNullString, txtCaptionOfDestWindow.Text)
    API_CopyMemory arrBytBuf(1), ByVal txtMsg.Text, Len(txtMsg.Text)
    
    cds.dwData = 54321  ' Our magic number, used like a password, set with property 'MsgHandle' .
    cds.cbData = Len(txtMsg.Text) + 1
    cds.lpData = VarPtr(arrBytBuf(1))
    
    API_SendMessage hWnd, WM_COPYDATA, Me.hWnd, cds
    
End Sub


' #*#
