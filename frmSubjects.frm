VERSION 5.00
Begin VB.Form frmSubjects 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subjects"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmSubjects.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "A&pply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtSubject 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Clo&se"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2760
      Width           =   1400
   End
   Begin VB.CommandButton cmdQuestions 
      Caption         =   "&Questions"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   1400
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   1400
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1400
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1400
   End
   Begin VB.ListBox lstSubject 
      Height          =   2985
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuQuestions 
         Caption         =   "Questions"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmSubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmSubjects
' Purpose   : For displaying, adding, editing and deleting subjects
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset

Private Sub cmdAdd_Click()

    Me.Height = 5000
    txtSubject.Text = ""
    txtSubject.SetFocus
    txtSubject.Tag = "Add"
    cmdApply.Enabled = False
    lstSubject.Enabled = False
    ToggleButtons False, False, False

End Sub

Private Sub cmdApply_Click()

  Dim tempRs As New Recordset
  Dim i As Long

    If Trim$(txtSubject.Text) = "" Then
        MsgBox "Please type a subject name", vbInformation, "Error Message"
        txtSubject.SetFocus
        Exit Sub
        'Check For Spaces
      ElseIf InStr(1, txtSubject.Text, " ") > 0 Then
        MsgBox "You have space in the subject name.", vbCritical + vbOKOnly, "Error Message"
        txtSubject.SetFocus
        Exit Sub
    End If
    OpenRecordSet tempRs, "tblSubjects", "SubjectName"
    If tempRs.RecordCount > 0 Then
        tempRs.MoveFirst
        For i = 1 To tempRs.RecordCount
            If tempRs!SubjectName = Trim$(txtSubject.Text) Then
                CloseRecordSet tempRs
                MsgBox "This subject already exist. Try again", vbInformation, "Error Message"
                cmdApply.Enabled = False
                txtSubject.SetFocus
                Exit Sub
            End If
            tempRs.MoveNext
        Next i
        CloseRecordSet tempRs
    End If
    If txtSubject.Tag = "Add" Then
        rs.AddNew
    End If
    rs!SubjectName = Trim$(txtSubject.Text)
    rs.update
    txtSubject.Tag = ""
    lstSubject.Enabled = True
    CloseRecordSet rs
    OpenRecordSet rs, "tblSubjects", , "SubjectName"
    LoadSubjects
    Me.Height = 4000

End Sub

Private Sub cmdCancel_Click()

  Dim hold As Long

    txtSubject.Tag = ""
    lstSubject.Enabled = True
    hold = lstSubject.ListIndex
    LoadSubjects
    lstSubject.ListIndex = hold
    If hold > 0 Then
        rs.AbsolutePosition = hold
    End If
    Me.Height = 4000

End Sub

Private Sub cmdClose_Click()

    CloseRecordSet rs
    Load frmMain
    frmMain.Show
    Unload Me

End Sub

Private Sub cmdDelete_Click()

  Dim tRs As New Recordset

    If lstSubject.ListIndex < 1 Then
        MsgBox "Select a Subject to delete"
        Exit Sub
      Else
        If MsgBox("Are you sure you want to delete" & vbNewLine & "the selected subject?" & "This will delete all" & vbNewLine & "questions and test scores under it" & vbNewLine & "You will have reset test options", vbYesNo, "Delete Subject") = vbYes Then
            OpenRecordSet tRs, "tblQuestion", , , "SubjectID = " & rs!SubjectID
            If tRs.RecordCount > 0 Then
                tRs.MoveFirst
                Do While Not tRs.EOF
                    tRs.delete
                    tRs.update
                    tRs.MoveNext
                Loop
            End If
            CloseRecordSet tRs
            OpenRecordSet tRs, "tblScores", , , "SubjectID = " & rs!SubjectID
            If tRs.RecordCount > 0 Then
                tRs.MoveFirst
                Do While Not tRs.EOF
                    tRs.delete
                    tRs.update
                    tRs.MoveNext
                Loop
            End If
            CloseRecordSet tRs

            OpenRecordSet tRs, "tblAdmin"
            If tRs.RecordCount > 0 Then
                tRs.MoveFirst
                tRs.delete
                tRs.update
            End If
            CloseRecordSet tRs

            rs.delete
            rs.update
            LoadSubjects
        End If
    End If

End Sub

Private Sub cmdEdit_Click()

    If lstSubject.ListIndex < 1 Then
        MsgBox "Select a subject from the list"
        Exit Sub
      Else
        Me.Height = 5000
        txtSubject.Text = lstSubject.Text
        txtSubject.SetFocus
        cmdApply.Enabled = False
        lstSubject.Enabled = False
        ToggleButtons False, False, False
        txtSubject.Tag = "Edit"
    End If

End Sub

Private Sub cmdQuestions_Click()

    If lstSubject.ListIndex > 0 Then
        frmQuestions.SubjectID = rs!SubjectID
      Else
        MsgBox "Select a subject from the list", vbCritical, "Error message"
        Exit Sub
    End If
    Load frmQuestions
    frmQuestions.Show
    Me.Hide

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    CustomTxtBox txtSubject, False, UpperCase
    Me.Height = 4000
    OpenRecordSet rs, "tblSubjects", , "SubjectName"
    LoadSubjects

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.picMenu.Enabled = True

End Sub

Private Sub LoadSubjects()

  Dim i As Long

    lstSubject.Clear
    If rs.RecordCount = 0 Then
        MsgBox "There are no registered subjects.", vbInformation, "Loading Subjects"
        ToggleButtons True, False, True
      Else
        rs.MoveFirst
        lstSubject.AddItem "Subjects List", 0
        For i = 1 To rs.RecordCount
            lstSubject.AddItem rs!SubjectName, rs.AbsolutePosition
            rs.MoveNext
        Next i
        rs.MoveFirst
        ToggleButtons True, True, True
    End If

End Sub

Private Sub lstSubject_Click()

    If lstSubject.ListIndex > 0 Then
        rs.AbsolutePosition = lstSubject.ListIndex
    End If

End Sub

Private Sub lstSubject_DblClick()

    If lstSubject.ListCount > 0 Then
        cmdQuestions.Value = True
    End If

End Sub

Private Sub lstSubject_ItemCheck(Item As Integer)

    If Item > 0 Then
        rs.AbsolutePosition = Item
    End If

End Sub

Private Sub lstSubject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If

End Sub

Private Sub mnuAdd_Click()

    cmdAdd.Value = True

End Sub

Private Sub mnuClose_Click()

    cmdClose.Value = True

End Sub

Private Sub mnuDelete_Click()

    cmdDelete.Value = True

End Sub

Private Sub mnuEdit_Click()

    cmdEdit.Value = True

End Sub

Private Sub mnuQuestions_Click()

    cmdQuestions.Value = True

End Sub

Private Sub SwitchButtons(bUp As Boolean)

    If bUp Then
        cmdQuestions.Default = True
        cmdClose.Cancel = True
      Else
        cmdApply.Default = True
        cmdCancel.Cancel = True
    End If

End Sub

Private Sub ToggleButtons(Add As Boolean, Edit As Boolean, Clse As Boolean)

    cmdAdd.Enabled = Add
    mnuAdd.Enabled = Add
    cmdEdit.Enabled = Edit
    mnuEdit.Enabled = Edit
    cmdDelete.Enabled = Edit
    mnuDelete.Enabled = Edit
    cmdQuestions.Enabled = Edit
    mnuQuestions.Enabled = Edit
    cmdClose.Enabled = Clse
    mnuClose.Enabled = Clse
    SwitchButtons Clse

End Sub

Private Sub txtSubject_Change()

    cmdApply.Enabled = True

End Sub

