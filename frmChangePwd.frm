VERSION 5.00
Begin VB.Form frmChangePwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   Icon            =   "frmChangePwd.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOldPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtNewPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtCNewPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change &Password"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblOldPass 
      BackStyle       =   0  'Transparent
      Caption         =   "&Old Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confir&m New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmChangePwd
' Purpose   : For changing user password
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdChange_Click()

    If Trim$(txtOldPass.Text) <> User.Password Then
        MsgBox "The old password you entered is not correct.", vbExclamation, "Change Password"
        Exit Sub
    End If
    If Trim$(txtNewPass.Text) = "" Then
        MsgBox "You haven't entered a new password.", vbExclamation, "Change Password"
        txtNewPass.SetFocus
        Exit Sub
        'Check For Spaces
      ElseIf InStr(1, txtNewPass.Text, " ") > 0 Then
        MsgBox "You have space in the new password", vbCritical + vbOKOnly, "Change Password"
        txtNewPass.SetFocus
        Exit Sub
      ElseIf Trim$(txtNewPass.Text) <> Trim$(txtCNewPass.Text) Then
        MsgBox "New password diffrent from the one you confirmed", vbExclamation, "Change Password"
        txtCNewPass.SetFocus
        Exit Sub
    End If
    OpenRecordSet rs, "tblUsers", , , "OfficialNo = " & User.OfficialNo
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        rs!Password = Trim$(txtNewPass.Text)
        rs.update
        User.Password = Trim$(txtNewPass.Text)
        MsgBox "You new password has been registered", vbInformation, "Change Password"
      Else
        MsgBox "Can't register the new password", vbCritical, "Change Password"
    End If
    CloseRecordSet rs
    Unload Me

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    CustomTxtBox txtOldPass, False, LowerCase
    CustomTxtBox txtNewPass, False, LowerCase
    CustomTxtBox txtCNewPass, False, LowerCase

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.picMenu.Enabled = True

End Sub

