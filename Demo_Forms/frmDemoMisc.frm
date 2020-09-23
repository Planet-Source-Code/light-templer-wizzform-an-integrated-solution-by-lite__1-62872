VERSION 5.00
Begin VB.Form frmDemoMisc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAC5AD&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "   ___ Demo *Misc* ___"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10440
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lbList 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      ItemData        =   "frmDemoMisc.frx":0000
      Left            =   180
      List            =   "frmDemoMisc.frx":0002
      TabIndex        =   0
      Top             =   1245
      Width           =   10065
   End
   Begin WizzFormDemo.ucWizzForm ucWizzForm 
      Left            =   210
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblHeadline 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo :   Here is a list with some more usefull subs/funcs exported for you by  WizzForm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1020
      TabIndex        =   1
      Top             =   285
      Width           =   8670
   End
End
Attribute VB_Name = "frmDemoMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Just some text infos in a listbox

Private Sub Form_Load()

    With lbList
        .AddItem " "
        .AddItem "EXPORTED FUNCTIONS (most of)"
        .AddItem "   ShowFormWithoutActivation()  [No focus / selected form change]"
        .AddItem "   PutFormOnTop()               [Works in W2K and above, too ... ;) ]"
        .AddItem "   CenterFormInWorkArea()       [Center form consider task / office bar]"
        .AddItem "   IsFunctionExported()         [Check a DLL for an API function]"
        .AddItem "   HiWord()                     [Higher 16 Bit part of a long value]"
        .AddItem "   LoWord()                     [Lower  16 Bit part of a long value]"
        .AddItem " "
        .AddItem "PROPERTIES (there are much more)"
        .AddItem "   FormMaxPosX                  [When max size is set and max button pressed ...]"
        .AddItem "   FormMaxPosY                  [When max size is set and max button pressed ...]"
        .AddItem "   FullDrag                     [Override system control panels setting]"
        .AddItem " "
        .AddItem " "
        .AddItem "Please have a closer look to the VERY fast 'DrawTopDownGradient()', too."
        .AddItem " "
        .AddItem "Read the comments at start of each sub to get further informations. Thx."
        
    End With
    
End Sub

' #*#
