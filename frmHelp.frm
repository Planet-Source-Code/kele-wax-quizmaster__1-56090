VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QUIZ MASTER HELP"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   240
      Width           =   8175
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Content"
      Begin VB.Menu mnuIntroduction 
         Caption         =   "Introduction"
      End
      Begin VB.Menu mnuRegUsers 
         Caption         =   "Registration of Users"
         Begin VB.Menu mnuRegTec 
            Caption         =   "Teachers/Administrators"
         End
         Begin VB.Menu mnuRegStu 
            Caption         =   "Students"
         End
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "Logging In"
      End
      Begin VB.Menu mnuTechers 
         Caption         =   "Teachers/Administrators"
         Begin VB.Menu mnuManUsers 
            Caption         =   "Manage Users"
         End
         Begin VB.Menu mnuQuizLig 
            Caption         =   "Manage Quiz Library"
         End
         Begin VB.Menu mnuQOptions 
            Caption         =   "Select Quiz Options"
         End
         Begin VB.Menu mnuPruneDB 
            Caption         =   "Prune Result DataBase"
         End
         Begin VB.Menu mnuCPwd 
            Caption         =   "Changing Password"
         End
         Begin VB.Menu mnuLOut 
            Caption         =   "Logging Out"
         End
         Begin VB.Menu mnuQuit 
            Caption         =   "Quitting"
         End
      End
      Begin VB.Menu mnuStudent 
         Caption         =   "Students"
         Begin VB.Menu mnuStartQ 
            Caption         =   "Start Quiz"
         End
         Begin VB.Menu mnuVResult 
            Caption         =   "View Result"
         End
         Begin VB.Menu mnuCPwd1 
            Caption         =   "Changing Password"
         End
         Begin VB.Menu mnuLogout1 
            Caption         =   "Logging Out"
         End
         Begin VB.Menu mnuQuit1 
            Caption         =   "Quitting"
         End
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHelp
' Purpose   : Offers a quick help on how to use application
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private Sub DocLocation(ByVal vNewValue As Double)

    txtHelp.SelStart = 1 ' force reset
    If vNewValue < 0 Then
        vNewValue = 0
    End If
    If vNewValue > 100 Then
        vNewValue = 100
    End If
    txtHelp.SelStart = CLng(vNewValue / 100 * Len(txtHelp.Text))
    DoEvents

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    OnTop Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.picMenu.Enabled = True
    frmMain.WindowState = 2

End Sub

Private Sub mnuClose_Click()

    OnTop Me, False
    OnTop frmMain
    Unload Me

End Sub

Private Sub mnuCPwd1_Click()

    DocLocation 98.624

End Sub

Private Sub mnuCPwd_Click()

    DocLocation 98.624

End Sub

Private Sub mnuIntroduction_Click()

    DocLocation 7.6

End Sub

Private Sub mnuLogin_Click()

    DocLocation 36.67

End Sub

Private Sub mnuLogout1_Click()

    DocLocation 99.92

End Sub

Private Sub mnuLOut_Click()

    DocLocation 99.92

End Sub

Private Sub mnuManUsers_Click()

    DocLocation 52.179

End Sub

Private Sub mnuPruneDB_Click()

    DocLocation 74.49

End Sub

Private Sub mnuQOptions_Click()

    DocLocation 70.63

End Sub

Private Sub mnuQuit1_Click()

    DocLocation 99.92

End Sub

Private Sub mnuQuit_Click()

    DocLocation 99.92

End Sub

Private Sub mnuQuizLig_Click()

    DocLocation 60.057

End Sub

Private Sub mnuRegStu_Click()

    DocLocation 33.887

End Sub

Private Sub mnuRegTec_Click()

    DocLocation 20

End Sub

Private Sub mnuStartQ_Click()

    DocLocation 85.209

End Sub

Private Sub mnuVResult_Click()

    DocLocation 94.765

End Sub

