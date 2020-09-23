VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mPopup 
      Caption         =   "POPUP"
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu hsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu hsep 
         Caption         =   "-"
      End
      Begin VB.Menu mSecurity 
         Caption         =   "Security"
         Begin VB.Menu mKiller 
            Caption         =   "Popup Killer"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mShow 
         Caption         =   "Show Window.."
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mAbout_Click()
FrmAbout.Show 1
End Sub

Private Sub mExit_Click()

If MsgBox("Are you sure you wish to quit?", vbInformation + vbYesNo, "Question") = vbYes Then
    frmMain.TIcon.RemoveIcon Me
    End
End If

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

Private Sub mShow_Click()

frmMain.WindowState = vbNormal
frmMain.Show

End Sub
