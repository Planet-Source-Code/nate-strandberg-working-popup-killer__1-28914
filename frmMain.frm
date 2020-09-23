VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Da'Zoom Popup Killer"
   ClientHeight    =   3375
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView vindows 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Location"
         Object.Width           =   5027
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "URL"
         Object.Width           =   4833
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Active Windows:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mSecurity 
      Caption         =   "&Security"
      Begin VB.Menu mKiller 
         Caption         =   "Popup Killer"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mRegister 
         Caption         =   "Register"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents IEWin As cIEWindows
Attribute IEWin.VB_VarHelpID = -1

Public TIcon As New clsTray

Private Sub Form_Load()

'Me.Hide
'frmNotice.Show 1

Active = True

Set IEWin = New cIEWindows
MakeTopMost Me.hwnd

TIcon.RemoveIcon Me
TIcon.ShowIcon Me
TIcon.ChangeToolTip Me, "Nate's Popup Killer"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If TIcon.bRunningInTray Then

    Select Case X
    
        Case 7725:
            Me.WindowState = vbNormal
            Me.Show
            
        Case 7755:
            PopupMenu frmMenu.mPopup

    End Select
    
End If

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If MsgBox("Are you sure you wish to quit?", vbInformation + vbYesNo, "Question") = vbYes Then
    TIcon.RemoveIcon Me
    End
Else
    Cancel = -1
End If

End Sub

Private Sub IEWin_IEWindowRegistered()
UpdateList
End Sub

Private Sub IEWin_IEWindowRevoked()
UpdateList
End Sub

Private Sub mAbout_Click()
FrmAbout.Show 1
End Sub

Private Sub mExit_Click()

If MsgBox("Are you sure you wish to quit?", vbInformation + vbYesNo, "Question") = vbYes Then
    TIcon.RemoveIcon Me
    End
End If

End Sub

Private Sub UpdateList()

On Error Resume Next

Dim temp As IE_Class
Dim itm As ListItem

vindows.ListItems.Clear

For Each temp In IEWin
   Set itm = vindows.ListItems.Add(, "K" & CStr(temp.IEHandle), temp.IEctl.LocationName)
   itm.SubItems(1) = temp.IEctl.LocationURL
Next temp

Set itm = Nothing

End Sub

Private Sub mPopup_Click()

End Sub

Private Sub mKiller_Click()

If mKiller.Checked = True Then
    mKiller.Checked = False
    Active = False
Else
    mKiller.Checked = True
    Active = True
End If

End Sub

Private Sub mRegister_Click()
frmRegister.Show 1
End Sub
