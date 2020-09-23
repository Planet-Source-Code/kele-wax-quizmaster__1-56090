VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUserInformation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   Icon            =   "frmStudentInfomation.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
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
      Left            =   6120
      TabIndex        =   13
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   8040
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudentInfomation.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudentInfomation.frx":0554
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "&View Selected Student's Result"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   4800
      Width           =   975
   End
   Begin MSComctlLib.TreeView tvwStudents 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8493
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      OLEDragMode     =   1
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   5160
      TabIndex        =   12
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   4200
      TabIndex        =   11
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3240
      TabIndex        =   16
      Top             =   360
      Width           =   5655
      Begin VB.CheckBox chkBar 
         Caption         =   "&Bar Selected User"
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
         Left            =   3000
         TabIndex        =   8
         ToolTipText     =   "This will prevent selected user from using this application"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.OptionButton optStudent 
         Caption         =   "&Student"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   2760
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optAdministrator 
         Caption         =   "Adminis&trator"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtOfficialNo 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   3255
      End
      Begin VB.ComboBox cboRank 
         Height          =   315
         ItemData        =   "frmStudentInfomation.frx":0666
         Left            =   1920
         List            =   "frmStudentInfomation.frx":068E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtInitials 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Official &Number :"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Surna&me :"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "&Initials :"
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
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "&Rank :"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmUserInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmUserInformation
' Purpose   : For displaying information on all register users. Can only be used by teachers
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset
Private bolSkip As Boolean

Private Sub ClearRecord()

    cboRank.ListIndex = 0
    txtInitials = ""
    txtSurname = ""
    txtOfficialNo = ""
    optStudent.Value = True
    chkBar.Value = False
    cmdResult.Visible = False

End Sub

Private Sub cmdAdd_Click()

    MsgBox "Fill all information and then click Update.", , "Add New User"
    ClearRecord
    LockUp False
    ToggleButtons False, False, False, True, False
    tvwStudents.Enabled = False
    cmdUpdate.Tag = "Add"

End Sub

Private Sub cmdCancel_Click()

    ShowRecord
    cmdUpdate.Tag = ""
    tvwStudents.Enabled = True
    LockUp True

End Sub

Private Sub cmdDelete_Click()

  Dim ToShow As Boolean
  Dim tRs As New Recordset

    ToShow = False
    cmdDelete.Tag = ""
    If MsgBox("Are you sure you want to delete " & rs!Rank & " " & rs!Initials & " " & rs!Surname, vbYesNo, "Delete User Information") = vbYes Then
        OpenRecordSet tRs, "tblScores", , , "OfficialNO = " & rs!OfficialNo
        If tRs.RecordCount > 0 Then
            tRs.MoveFirst
            Do While Not tRs.EOF
                tRs.delete
                tRs.update
                tRs.MoveNext
            Loop
        End If
        CloseRecordSet tRs
        rs.delete
        rs.update
        CloseRecordSet rs
        OpenRecordSet rs, "tblUsers", , "Student, Rank, OfficialNo", "OfficialNo <> " & User.OfficialNo
        FillTvw
        LockUp True
        ClearRecord
        ToggleButtons True, False, False, False, True
    End If

End Sub

Private Sub cmdEdit_Click()

    MsgBox "Make your changes and then click Update.", , "Edit User Information"
    LockUp False
    ToggleButtons False, False, False, True, False
    tvwStudents.Enabled = False
    cmdUpdate.Tag = "Edit"

End Sub

Private Sub cmdExit_Click()

    CloseRecordSet rs
    Load frmMain
    Unload Me

End Sub

Private Sub cmdResult_Click()

    Load frmResult
    If CanContinue(frmResult) Then
        Me.Hide
    End If

End Sub

Private Sub cmdUpdate_Click()

  Dim tempRs As New Recordset
  Dim i As Long

    If cboRank.Text = "" Then
        MsgBox "Select a rank", vbCritical, "Error Message"
        cboRank.SetFocus
        Exit Sub
      ElseIf Trim$(txtInitials.Text) = "" Then
        MsgBox "Enter the users initials", vbCritical, "Error Message"
        txtInitials.SetFocus
        Exit Sub
      ElseIf Trim$(txtSurname.Text) = "" Then
        MsgBox "Enter the users surnamae.", vbCritical, "Error Message"
        txtSurname.SetFocus
        Exit Sub
      ElseIf Trim$(txtOfficialNo.Text) = "" Then
        MsgBox "Enter the users official number.", vbCritical, "Error Message"
        txtOfficialNo.SetFocus
        Exit Sub
        'Check For Spaces
      ElseIf InStr(1, txtInitials.Text, " ") > 0 Then
        MsgBox "You have space in the initials", vbCritical + vbOKOnly, "Error Message"
        txtInitials.SetFocus
        Exit Sub
        'Check For Spaces
      ElseIf InStr(1, txtSurname.Text, " ") > 0 Then
        MsgBox "You have space in the surname", vbCritical + vbOKOnly, "Error Message"
        txtSurname.SetFocus
        Exit Sub
        'Check For Spaces
      ElseIf InStr(1, txtOfficialNo.Text, " ") > 0 Then
        MsgBox "You have space in the official number", vbCritical + vbOKOnly, "Error Message"
        txtOfficialNo.SetFocus
        Exit Sub
    End If

    OpenRecordSet tempRs, "tblUsers", , "Student, Rank, OfficialNo"
    Select Case cmdUpdate.Tag
      Case "Add"
        tempRs.MoveFirst
        For i = 1 To tempRs.RecordCount
            If txtOfficialNo.Text = tempRs!OfficialNo Then
                MsgBox "A user with official number " & tempRs!OfficialNo & " already exist.", , "Error adding new user"
                txtOfficialNo = ""
                txtOfficialNo.SetFocus
                CloseRecordSet tempRs
                Exit Sub
            End If
            tempRs.MoveNext
        Next i
        CloseRecordSet tempRs
        rs.AddNew
        rs!Password = Trim$(txtOfficialNo)
        GoTo Final
      Case "Edit"
        If CLng(Trim$(txtOfficialNo)) <> rs!OfficialNo Then
            tempRs.MoveFirst
            For i = 1 To tempRs.RecordCount
                If txtOfficialNo.Text = tempRs!OfficialNo Then
                    MsgBox "A user with official number " & tempRs!OfficialNo & " already exist.", , "Error adding new user"
                    txtOfficialNo = ""
                    txtOfficialNo.SetFocus
                    CloseRecordSet tempRs
                    Exit Sub
                End If
                tempRs.MoveNext
            Next i
            CloseRecordSet tempRs
            GoTo Final
        End If

    End Select
Final:
    rs!Rank = cboRank.Text
    rs!Initials = Trim$(txtInitials)
    rs!Surname = Trim$(txtSurname)
    rs!OfficialNo = CLng(Trim$(txtOfficialNo))
    rs!Student = optStudent.Value
    rs!Active = Not CBool(CInt(chkBar.Value) * -1)
    rs.update
    CloseRecordSet rs
    OpenRecordSet rs, "tblUsers", , "Student, Rank, OfficialNo", "OfficialNo <> " & User.OfficialNo
    tvwStudents.Enabled = True
    FillTvw
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        If rs!OfficialNo = CLng(Trim$(txtOfficialNo)) Then
            Exit For
        End If
        rs.MoveNext
    Next i
    For i = 1 To tvwStudents.Nodes.Count
        If tvwStudents.Nodes(i).Tag = "N" & rs.AbsolutePosition Then
            tvwStudents.Nodes(i).Selected = True
            Exit For
        End If
    Next i
    LockUp True
    cmdResult.Visible = optStudent.Value
    ToggleButtons True, True, True, False, True

End Sub

Private Sub FillTvw()

  Dim i As Long
  Dim t As Long
  Dim xNode As Node

    tvwStudents.Nodes.Clear
    tvwStudents.Nodes.Add , , "Admin", "ADMINISTRATORS", 1, 2
    tvwStudents.Nodes.Add , , "Students", "STUDENTS", 1, 2
    For i = 0 To cboRank.ListCount - 1
        Set xNode = tvwStudents.Nodes.Add("Admin", tvwChild, cboRank.List(i) & "A", cboRank.List(i), 1, 2)
        xNode.EnsureVisible
        If rs.RecordCount < 1 Then
            GoTo Skip1
        End If
        rs.MoveFirst
        For t = 1 To rs.RecordCount
            If xNode.key = rs!Rank & "A" And Not rs!Student Then
                tvwStudents.Nodes.Add xNode.key, tvwChild, "N" & t, rs!Initials & " " & rs!Surname
            End If
            rs.MoveNext
        Next t
Skip1:
        Set xNode = tvwStudents.Nodes.Add("Students", tvwChild, cboRank.List(i) & "S", cboRank.List(i), 1, 2)
        xNode.EnsureVisible
        If rs.RecordCount < 1 Then
            GoTo Skip2
        End If
        rs.MoveFirst
        For t = 1 To rs.RecordCount
            If xNode.key = rs!Rank & "S" And rs!Student Then
                tvwStudents.Nodes.Add xNode.key, tvwChild, "N" & t, rs!Initials & " " & rs!Surname
            End If
            rs.MoveNext
        Next t
Skip2:
    Next i

    ToggleButtons True, False, False, False, True
    bolSkip = True

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    CustomTxtBox txtInitials, False, UpperCase
    CustomTxtBox txtSurname, False, UpperCase
    CustomTxtBox txtOfficialNo, True
    OpenRecordSet rs, "tblUsers", , "Student, Rank, OfficialNo", "OfficialNo <> " & User.OfficialNo
    FillTvw
    ClearRecord
    LockUp True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.picMenu.Enabled = True

End Sub

Private Sub LockUp(opt As Boolean)

    cboRank.Enabled = Not opt
    txtInitials.Locked = opt
    txtSurname.Locked = opt
    txtOfficialNo.Locked = opt
    optAdministrator.Enabled = Not opt
    optStudent.Enabled = Not opt
    chkBar.Enabled = Not opt

End Sub

Private Sub ShowRecord()

    If bolSkip Then
        Exit Sub
    End If
    cboRank.Text = rs!Rank
    SelStudent.Rank = rs!Rank
    txtInitials.Text = rs!Initials
    SelStudent.Initials = rs!Initials
    txtSurname.Text = rs!Surname
    SelStudent.Surname = rs!Surname
    txtOfficialNo.Text = rs!OfficialNo
    SelStudent.OfficialNo = rs!OfficialNo
    optStudent.Value = rs!Student
    optAdministrator.Value = Not rs!Student
    chkBar.Value = CInt(Not rs!Active) * -1
    cmdResult.Visible = rs!Student

    ToggleButtons True, True, True, False, True

End Sub

Private Sub ToggleButtons(Add As Boolean, Edit As Boolean, delete As Boolean, update As Boolean, cExit As Boolean)

    cmdAdd.Enabled = Add
    cmdEdit.Enabled = Edit
    cmdDelete.Enabled = delete
    cmdCancel.Enabled = update
    cmdUpdate.Enabled = update
    cmdExit.Enabled = cExit

End Sub

Private Sub tvwStudents_NodeClick(ByVal Node As MSComctlLib.Node)

    If Left$(Node.key, 1) = "N" Then
        rs.AbsolutePosition = Val(Mid$(Node.key, 2))
        bolSkip = False
        ShowRecord
        ToggleButtons True, True, True, False, True
      Else
        ClearRecord
        ToggleButtons True, False, False, False, True
        bolSkip = True
    End If
    LockUp True

End Sub

