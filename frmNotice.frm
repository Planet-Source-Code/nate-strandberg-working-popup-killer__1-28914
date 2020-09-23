VERSION 5.00
Begin VB.Form frmNotice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer lie 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label tCount 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Time Remaining:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Evaluation Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Da'Zoom Popup Killer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TimeTil As Long
Private Current As Long

Private Sub cmdOkay_Click()

Unload Me
frmMain.Show

End Sub

Private Sub Form_Load()

Current = 5
Active = True

End Sub

Private Sub lie_Timer()

TimeTil = TimeTil + 1
Current = Current - 1

tCount = Current

If TimeTil >= 5 Then
    Current = 0
    lie.Enabled = False
    cmdOkay.Enabled = True
    tCount = "Software enabled. Please Register!"
End If

End Sub

