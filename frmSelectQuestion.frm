VERSION 5.00
Begin VB.Form frmSelectQuestion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Questions For Test"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   Icon            =   "frmSelectQuestion.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   26
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkEqualMarks 
      Caption         =   "All Questions Carry &Equal Marks"
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame fmMarks 
      Caption         =   "Marks Distribution  --  (%)"
      Height          =   1815
      Left            =   6480
      TabIndex        =   31
      Top             =   720
      Width           =   2775
      Begin VB.TextBox txtWMark 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtTMark 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtMMark 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Wr&itten :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "&True or False :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Multiple &Choice :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame fmAdvanced 
      Caption         =   "Frame4"
      Height          =   1455
      Left            =   3120
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Frame fmStartQNo 
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2895
         Begin VB.ComboBox cboQuestionNo 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "St&art From Question No :"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkQuestionNo 
         Caption         =   "As&k Questions Randomly"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame fmTime 
      Caption         =   "Test Duration"
      Height          =   855
      Left            =   3120
      TabIndex        =   28
      Top             =   720
      Width           =   3135
      Begin VB.TextBox txtMinute 
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "00"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtHour 
         Height          =   285
         Left            =   720
         MaxLength       =   1
         TabIndex        =   17
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Mi&nutes :"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "&Hours :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkDistribution 
      Caption         =   "&Random Selection"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame fmQDist 
      Caption         =   "Question Distribution"
      Height          =   2535
      Left            =   360
      TabIndex        =   27
      Top             =   720
      Width           =   2535
      Begin VB.TextBox txtWritten 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Text            =   " "
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtBoolean 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtMCQ 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "&Written :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "True or &False :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "&Multiple Choice :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox txtTotalQ 
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.ComboBox cboSubject 
      Height          =   315
      ItemData        =   "frmSelectQuestion.frx":0442
      Left            =   1320
      List            =   "frmSelectQuestion.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Total Number of &Questions :"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "&Subject :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmSelectQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmSelectQuestion
' Purpose   : Used to set test options
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset
Private NoSubject As Boolean
Private StartQType(1 To 3) As Long
Private LastQType As Integer

Private Function AllQuestCount() As Long

  Dim rsSubject As New Recordset
  Dim rsQuest As New Recordset

    If NoSubject Then
        AllQuestCount = 0
      Else
        If cboSubject.ListIndex > 0 Then
            OpenRecordSet rsSubject, "tblSubjects", , , "SubjectName = '" & Trim$(cboSubject.Text) & "'"
            OpenRecordSet rsQuest, "tblQuestion", , , "SubjectID = " & rsSubject!SubjectID
            AllQuestCount = rsQuest.RecordCount
            CloseRecordSet rsSubject
            CloseRecordSet rsQuest
          Else
            AllQuestCount = 0
        End If
    End If

End Function

Private Sub cboSubject_Change()

    txtTotalQ_Change

End Sub

Private Sub cboSubject_Click()

    txtTotalQ_Change

End Sub

Private Sub chkDistribution_Click()

    RanDist

End Sub

Private Sub chkEqualMarks_Click()

    EqualMk

End Sub

Private Sub chkQuestionNo_Click()

    SetStartQNo

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

  Dim rsSubject As New Recordset

    If Val(txtTotalQ.Text) = 0 Then
        If MsgBox("The total number of questions to be asked has been set to zero." & vbNewLine & "This means that no test can be taken by any student." & vbNewLine & "Do you wish to continue?", vbYesNo, "Select questions for test") = vbNo Then
            txtTotalQ.SetFocus
            Exit Sub
        End If

      ElseIf chkDistribution.Value = 0 Then
        If Val(txtMCQ.Text) + Val(txtBoolean.Text) + Val(txtWritten.Text) < Val(txtTotalQ.Text) Then
            MsgBox "Please check distribution of question among question types." & vbNewLine & "Distribution is less than total number of questions.", , "Select questions for test"
            txtMCQ.SetFocus
            Exit Sub
          ElseIf Val(txtMCQ.Text) + Val(txtBoolean.Text) + Val(txtWritten.Text) > Val(txtTotalQ.Text) Then
            MsgBox "Please check distribution of question among question types." & vbNewLine & "Distribution is more than total number of questions.", , "Select questions for test"
            txtMCQ.SetFocus
            Exit Sub
        End If
      ElseIf chkDistribution.Value Then
        txtMCQ.Text = "0"
        txtBoolean.Text = "0"
        txtWritten.Text = "0"
        StartQType(1) = 0
        StartQType(2) = 0
        StartQType(3) = 0
    End If

    If (Val(txtHour.Text) * 60) + Val(txtMinute.Text) = 0 Then
        If MsgBox("You have not set any time for the test duration." & vbNewLine & "This means that no test can be taken by any student." & vbNewLine & "Do you wish to continue?", vbYesNo, "Select questions for test") = vbNo Then
            txtMinute.SetFocus
            Exit Sub
        End If
    End If

    If chkEqualMarks.Value = 0 Then
        If Val(txtMMark.Text) + Val(txtTMark.Text) + Val(txtWMark.Text) > 100 Then
            MsgBox "Please check distribution of marks among question types." & vbNewLine & "Distribution is more 100%.", , "Select questions for test"
            txtMMark.SetFocus
            Exit Sub
          ElseIf Val(txtMMark.Text) + Val(txtTMark.Text) + Val(txtWMark.Text) < 100 Then
            MsgBox "Please check distribution of question among question types." & vbNewLine & "Distribution is less 100%.", , "Select questions for test"
            txtMMark.SetFocus
            Exit Sub
        End If
      ElseIf chkEqualMarks.Value Then
        txtMMark.Text = "0"
        txtTMark.Text = "0"
        txtWMark.Text = "0"
    End If

    If rs.RecordCount = 0 Then
        rs.AddNew
      Else
        rs.MoveFirst
    End If
    If cboSubject.ListIndex > 0 Then
        OpenRecordSet rsSubject, "tblSubjects", , , "SubjectName = '" & cboSubject.Text & "'"
        rs!SubjectID = rsSubject!SubjectID
        CloseRecordSet rsSubject
      Else
        rs!SubjectID = 0
    End If

    rs!TotalNumber = Val(txtTotalQ.Text)
    rs!Ran_Dist = chkDistribution.Value
    rs!MCQNo = Val(txtMCQ.Text)
    rs!TrueFalseNo = Val(txtBoolean.Text)
    rs!WrittenNo = Val(txtWritten.Text)
    rs!Duration = (Val(txtHour.Text) * 60) + Val(txtMinute.Text)
    rs!StartMCQ = StartQType(1)
    rs!StartTrueFalse = StartQType(2)
    rs!StartWritten = StartQType(3)
    rs!Equal = chkEqualMarks.Value
    rs!MCQ = Val(txtMMark.Text)
    rs!TrueFalse = Val(txtTMark.Text)
    rs!Written = Val(txtWMark.Text)
    rs.update
    Unload Me

End Sub

Private Sub CTBx()

  Dim oControl As Control

    For Each oControl In Me.Controls
        If TypeName(oControl) = "TextBox" Then
            CustomTxtBox oControl, True
        End If
    Next oControl

End Sub

Private Sub EqualMk()

    If chkEqualMarks.Value Then
        txtMMark.Text = "0"
        txtTMark.Text = "0"
        txtWMark.Text = "0"
    End If
    fmMarks.Enabled = Not CBool(CInt(chkEqualMarks.Value) * -1)

End Sub

Private Sub FillQType(Qtype As Integer)

  Dim i As Long

    SaveQTypeChanges
    cboQuestionNo.Clear
    cboQuestionNo.AddItem "0", 0
    If QuestCount(Qtype) > 0 Then
        For i = 1 To QuestCount(Qtype)
            cboQuestionNo.AddItem CStr(i), i
        Next i
    End If

End Sub

Private Sub FillSubjects()

  Dim tRs As New Recordset
  Dim i As Long

    OpenRecordSet tRs, "tblSubjects", , "SubjectName"
    cboSubject.Clear
    If tRs.RecordCount > 0 Then
        NoSubject = False
        cboSubject.AddItem "All Subjects", 0
        tRs.MoveFirst
        For i = 1 To tRs.RecordCount
            cboSubject.AddItem tRs!SubjectName, i
            tRs.MoveNext
        Next i
      Else
        NoSubject = True
        cboSubject.AddItem "No Subjects"
    End If
    CloseRecordSet tRs

End Sub

Private Sub Form_Load()

  Dim rsSubject As New Recordset

    DisableX Me.hwnd
    CTBx
    OpenRecordSet rs, "tblAdmin"
    FillSubjects
    If rs.RecordCount > 0 Then
        If rs!SubjectID > 0 Then
            OpenRecordSet rsSubject, "tblSubjects", , , "SubjectID = " & rs!SubjectID
            If rsSubject.RecordCount > 0 Then
                cboSubject.Text = rsSubject!SubjectName
              Else
                cboSubject.ListIndex = 0
            End If
            CloseRecordSet rsSubject
          Else
            cboSubject.ListIndex = 0
        End If

        txtTotalQ.Text = rs!TotalNumber
        chkDistribution.Value = CInt(rs!Ran_Dist) * -1
        txtMCQ.Text = rs!MCQNo
        txtBoolean.Text = rs!TrueFalseNo
        txtWritten.Text = rs!WrittenNo
        txtHour.Text = rs!Duration \ 60
        txtMinute.Text = rs!Duration Mod 60
        StartQType(1) = rs!StartMCQ
        StartQType(2) = rs!StartTrueFalse
        StartQType(3) = rs!StartWritten
        chkEqualMarks.Value = CInt(rs!Equal) * -1
        txtMMark.Text = rs!MCQ
        txtTMark.Text = rs!TrueFalse
        txtWMark.Text = rs!Written
        RanDist
        EqualMk
      Else
        cboSubject.ListIndex = 0
        txtTotalQ.Text = 0
        chkDistribution.Value = 1
        txtHour.Text = "0"
        txtMinute.Text = "00"
        StartQType(1) = 0
        StartQType(2) = 0
        StartQType(3) = 0
        chkEqualMarks.Value = 1
        RanDist
        EqualMk
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseRecordSet rs
    frmMain.picMenu.Enabled = True

End Sub

Private Function QuestCount(Qtype As Integer) As Long

  Dim rsQuest As New Recordset
  Dim rsSubject As New Recordset

    QuestCount = 0
    If (Qtype < 1) Or (Qtype > 3) Then
        Exit Function
    End If

    If NoSubject Then

      ElseIf AllQuestCount = 0 Then

      Else
        If cboSubject.ListIndex > 0 Then
            OpenRecordSet rsSubject, "tblSubjects", , , "SubjectName = '" & Trim$(cboSubject.Text) & "'"
            OpenRecordSet rsQuest, "tblQuestion", , , "SubjectID = " & rsSubject!SubjectID & "AND QType = " & Qtype
            QuestCount = rsQuest.RecordCount
            CloseRecordSet rsSubject
            CloseRecordSet rsQuest
          Else
            OpenRecordSet rsQuest, "tblQuestion", , , "QType = " & Qtype
            QuestCount = rsQuest.RecordCount
            CloseRecordSet rsQuest
        End If
    End If

End Function

Private Sub RanDist()

    If chkDistribution.Value Then
        txtMCQ.Text = "0"
        txtBoolean.Text = "0"
        txtWritten.Text = "0"
        txtWritten.Locked = True
        txtBoolean.Locked = True
        txtMCQ.Locked = True
        LastQType = 0
        ShowQType
      Else
        txtMCQ.Locked = False
        txtBoolean.Locked = False
        txtWritten.Locked = False
    End If

End Sub

Private Sub SaveQTypeChanges()

    If LastQType Then
        StartQType(LastQType) = cboQuestionNo.ListIndex
    End If

End Sub

Private Sub SetStartQNo()

    fmStartQNo.Visible = Not CBool(CInt(chkQuestionNo.Value) * -1)
    If chkQuestionNo.Value Then
        cboQuestionNo.Text = 0
    End If

End Sub

Private Sub SetUp()

    If Val(txtTotalQ.Text) = 0 Then
        chkDistribution.Value = False
        txtMCQ.Text = ""
        txtBoolean.Text = ""
        txtWritten.Text = ""
        fmQDist.Enabled = False
      Else
        fmQDist.Enabled = True
    End If

End Sub

Private Sub ShowQType(Optional Qtype As Integer = 0)

    If Not chkDistribution.Value Then
        Select Case Qtype
          Case 0
            fmAdvanced.Visible = False
            LastQType = 0

          Case 1
            fmAdvanced.Caption = "Multiple Choice Questions"
            LastQType = 1
            fmAdvanced.Visible = Val(txtMCQ.Text)

          Case 2
            fmAdvanced.Caption = "True or False Questions"
            LastQType = 2
            fmAdvanced.Visible = Val(txtBoolean.Text)
          Case 3
            fmAdvanced.Caption = "Written Questions"
            LastQType = 3
            fmAdvanced.Visible = Val(txtWritten.Text)
        End Select
        SetStartQNo
      Else
        LastQType = 0
        fmAdvanced.Visible = False
    End If

End Sub

Private Sub txtBoolean_Change()

    If Val(txtBoolean.Text) > QuestCount(2) Then
        txtBoolean.Text = QuestCount(2)
      Else
        txtBoolean.Text = Val(txtBoolean.Text)
    End If
    ShowQType 2

End Sub

Private Sub txtBoolean_DblClick()

    If chkDistribution.Value = 0 Then
        txtBoolean.Text = QuestCount(2)
    End If

End Sub

Private Sub txtBoolean_GotFocus()

    FillQType 2
    ShowQType 2
    chkQuestionNo.Value = CInt(Not CBool(StartQType(2))) * -1
    cboQuestionNo.ListIndex = StartQType(2)

End Sub

Private Sub txtMCQ_Change()

    If Val(txtMCQ.Text) > QuestCount(1) Then
        txtMCQ.Text = QuestCount(1)
      Else
        txtMCQ.Text = Val(txtMCQ.Text)
    End If
    ShowQType 1

End Sub

Private Sub txtMCQ_DblClick()

    If chkDistribution.Value = 0 Then
        txtMCQ.Text = QuestCount(1)
    End If

End Sub

Private Sub txtMCQ_GotFocus()

    FillQType 1
    ShowQType 1
    chkQuestionNo.Value = CInt(Not CBool(StartQType(1))) * -1
    cboQuestionNo.ListIndex = StartQType(1)

End Sub

Private Sub txtMMark_Change()

    If Val(txtMCQ.Text) = 0 Then
        txtMMark.Text = "0"
    End If

End Sub

Private Sub txtTMark_Change()

    If Val(txtBoolean.Text) = 0 Then
        txtTMark.Text = "0"
    End If

End Sub

Private Sub txtTotalQ_Change()

    If Val(txtTotalQ.Text) > AllQuestCount Then
        txtTotalQ.Text = AllQuestCount
      Else
        txtTotalQ.Text = Val(txtTotalQ.Text)
    End If
    SetUp

End Sub

Private Sub txtTotalQ_DblClick()

    txtTotalQ.Text = AllQuestCount

End Sub

Private Sub txtWMark_Change()

    If Val(txtWritten.Text) = 0 Then
        txtWMark.Text = "0"
    End If

End Sub

Private Sub txtWritten_Change()

    If Val(txtWritten.Text) > QuestCount(3) Then
        txtWritten.Text = QuestCount(3)
      Else
        txtWritten.Text = Val(txtWritten.Text)
    End If
    ShowQType 3

End Sub

Private Sub txtWritten_DblClick()

    If chkDistribution.Value = 0 Then
        txtWritten.Text = QuestCount(3)
    End If

End Sub

Private Sub txtWritten_GotFocus()

    FillQType 3
    ShowQType 3
    chkQuestionNo.Value = CInt(Not CBool(StartQType(3))) * -1
    cboQuestionNo.ListIndex = StartQType(3)

End Sub

