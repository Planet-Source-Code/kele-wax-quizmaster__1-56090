VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4215
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3957.657
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOfficialNo 
      Height          =   345
      Left            =   1890
      TabIndex        =   1
      Top             =   135
      Width           =   2205
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Height          =   390
      Left            =   615
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2220
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2205
   End
   Begin VB.Label lblLabels 
      Caption         =   "Official &Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1680
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmLogin
' Purpose   : For logging into programme
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset

Private Sub cmdOK_Click()

    If Trim$(txtOfficialNo.Text) = "" Then
        MsgBox "Enter Your Official Number", , "Login Error"
        txtOfficialNo.SetFocus
        Exit Sub
      ElseIf Trim$(txtPassword.Text) = "" Then
        MsgBox "Enter Your Password", , "Login Error"
        txtPassword.SetFocus
        Exit Sub
    End If
    OpenRecordSet rs, "tblUsers", , , "OfficialNo = " & Val(txtOfficialNo.Text) & " And Password = '" & txtPassword.Text & "'"
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        With rs
            If !Active = False Then
                MsgBox "Sorry " & rs!Rank & " " & rs!Initials & " " & rs!Surname & vbNewLine & "Your access has been denied." & vbNewLine & "Please contact the administrator", vbCritical, "Error message"
                txtOfficialNo.Text = ""
                txtPassword.Text = ""
                txtOfficialNo.SetFocus
                Exit Sub
            End If

            User.Initials = !Initials
            User.OfficialNo = !OfficialNo
            User.Password = !Password
            User.Rank = !Rank
            User.Student = !Student
            User.Surname = !Surname
        End With
        CloseRecordSet rs
        Load frmMain
        frmMain.Show
        Unload Me
      Else
        MsgBox "Wrong Official Number or Password," & vbNewLine & "Please try again.", , "Error Logging in..."
        txtOfficialNo.Text = ""
        txtPassword.Text = ""
        txtOfficialNo.SetFocus
    End If

End Sub

Private Sub cmdQuit_Click()

    CloseRecordSet rs
    Unload Me

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    CustomTxtBox txtOfficialNo, True
    CustomTxtBox txtPassword, False, LowerCase
    CreateNewDataBase
    OpenRecordSet rs, "tblUsers", , , "Student = False"
    If rs.RecordCount > 0 Then
        CloseRecordSet rs
      Else
        CloseRecordSet rs
        If MsgBox("There is no record of an administrator in the database." & vbNewLine & "Do you want to registar as an administrator?", vbYesNo, "Login...") = vbYes Then
            frmRegAdmin.Show
        End If
        Unload Me
    End If

End Sub

