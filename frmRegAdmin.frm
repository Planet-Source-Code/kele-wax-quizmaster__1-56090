VERSION 5.00
Begin VB.Form frmRegAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register as an administrator"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "frmRegAdmin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3480
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   5295
      Begin VB.TextBox txtBox 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
      End
      Begin VB.ComboBox cboRank 
         Height          =   315
         ItemData        =   "frmRegAdmin.frx":0442
         Left            =   1800
         List            =   "frmRegAdmin.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtBox 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Confir&m Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ran&k :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "&Initials :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "&Last Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Official &Number :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmRegAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmRegAdmin
' Purpose   : For registring an Administrator when programme is first run
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdRegister_Click()

  Dim i As Integer

    If cboRank.Text = "" Then
        MsgBox "Select Your Rank", , "Register Administrator"
        Exit Sub
    End If
    For i = 0 To 4
        If txtBox(i) = "" Then
            MsgBox "Please fill in all parameters"
            txtBox(i).SetFocus
            Exit Sub
        End If
    Next i

    'Check For Spaces
    For i = 0 To 4
        If InStr(1, txtBox(i).Text, " ") > 0 Then
            MsgBox "Space in-between text, not allowed.", vbCritical + vbOKOnly, "Error Message"
            txtBox(i).SetFocus
            Exit Sub
        End If
    Next i

    If txtBox(3) <> txtBox(4) Then
        MsgBox "The Password you entered is diffrent from the one you confirmed."
        txtBox(3) = ""
        txtBox(4) = ""
        txtBox(3).SetFocus
        Exit Sub
    End If

    OpenRecordSet rs, "tblUsers", , , "OfficialNo = " & Val(txtBox(2))
    If rs.RecordCount = 0 Then
        CloseRecordSet rs
      Else
        CloseRecordSet rs
        MsgBox "The Official number you entered was registered before." & vbNewLine & "Try again", , "Error Registring Administrator"
        txtBox(3) = ""
        txtBox(3).SetFocus
        Exit Sub
    End If

    OpenRecordSet rs, "tblUsers"
    rs.AddNew

    rs!Rank = cboRank.Text
    rs!Initials = Trim$(txtBox(0))
    rs!Surname = Trim$(txtBox(1))
    rs!OfficialNo = Val(txtBox(2))
    rs!Password = Trim$(txtBox(3))
    rs!Student = False
    rs!Active = True
    rs.UpdateBatch

    If MsgBox("Congratulations " & cboRank.Text & " " & Trim$(txtBox(0)) & " " & Trim$(txtBox(1)) & vbNewLine & "You have been successfully registered as an administrator." & vbNewLine & "Do you want to login?", vbYesNo, "Register an administrator...") = vbYes Then

        User.Initials = rs!Initials
        User.OfficialNo = rs!OfficialNo
        User.Password = rs!Password
        User.Rank = rs!Rank
        User.Student = rs!Student
        User.Surname = rs!Surname

        CloseRecordSet rs
        Load frmMain
        frmMain.Show
        Unload Me
      Else
        CloseRecordSet rs
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    CustomTxtBox txtBox(0), False, UpperCase
    CustomTxtBox txtBox(1), False, UpperCase
    CustomTxtBox txtBox(2), True
    CustomTxtBox txtBox(3), False, LowerCase
    CustomTxtBox txtBox(4), False, LowerCase

End Sub

