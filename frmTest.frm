VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12915
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   Begin prjQuizMaster.ZoomPicCtl picQuestion 
      Height          =   5295
      Left            =   8160
      TabIndex        =   28
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9340
      BackColor       =   -2147483633
      AllowZoomIn     =   -1  'True
      AllowZoomOut    =   -1  'True
      UseQuickBar     =   -1  'True
   End
   Begin VB.Frame fmAns 
      Height          =   3255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   7695
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   3
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2400
         Width           =   6495
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "&D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2400
         Width           =   600
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   2
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1680
         Width           =   6495
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "&C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   1
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   960
         Width           =   6495
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "&B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   0
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   6495
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "&A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Timer tmrCountDown 
      Left            =   9000
      Top             =   4440
   End
   Begin VB.Frame fmNavigate 
      Height          =   855
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   7695
      Begin VB.CommandButton cmdNavigate 
         Cancel          =   -1  'True
         Caption         =   "&QUIT"
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
         Index           =   2
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&NEXT"
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
         Index           =   1
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&SKIP"
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
         Index           =   0
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fmQuestion 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      Begin VB.PictureBox picTime 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H008080FF&
         Height          =   495
         Left            =   6960
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   27
         Top             =   160
         Width           =   615
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   495
            Left            =   0
            Shape           =   3  'Circle
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtQuestion 
         Height          =   1095
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label lblRQuest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   6240
         TabIndex        =   7
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblTQuest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblQuestNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5640
         Picture         =   "frmTest.frx":0442
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Total Questions :"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Remaining Question :"
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Question No :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fmAns 
      Height          =   855
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   7695
      Begin VB.OptionButton optBoolean 
         Caption         =   "   &FALSE"
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
         Index           =   1
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optBoolean 
         Caption         =   "    &TRUE"
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
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fmAns 
      Height          =   1455
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   2400
      Width           =   7695
      Begin VB.TextBox txtAnswer 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTest
' Purpose   : For taking a test as specified by the test options set by an administrator
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset
Private Duration As Date
Private ElapsedMin As Long
Private QNo As Long
Private mStream As New ADODB.Stream
Private Fpath As String
Private UOption As Long
Private TimePie As New clsPieBar

Private Sub AddQuestion(Qtype As Integer)

    With Questions(QNo)
        .Question_No = rs.AbsolutePosition
        .Question_Type = Qtype
        .IsAnswered = False
        Select Case Qtype
          Case 1
            .Correct_Option = rs!CorrectOption
          Case 2
            .Correct_Choice = rs!Boolean
          Case 3
            .Correct_Answer = rs!Answer
        End Select
    End With

End Sub

Private Sub ClearAll()

  Dim i As Integer

    txtQuestion.Text = ""
    For i = 0 To 3
        txtOption(i).Text = ""
        cmdOption(i).BackColor = vbButtonFace
        cmdOption(i).Tag = ""
    Next i
    txtAnswer.Text = ""
    optBoolean(0).Value = False
    optBoolean(1).Value = False
    'picQuestion.Picture = LoadPicture
    picQuestion.UnloadImage
    lblQuestNo.Caption = ""
    lblRQuest.Caption = ""
    lblTQuest.Caption = ""

End Sub

Private Sub cmdNavigate_Click(Index As Integer)

    Select Case Index
      Case 0 'Skip
        Questions(QNo).IsAnswered = False
        If QNo >= Test.TotalNumber Then
            GoTo Skip
        End If
      Case 1 'Next
        If UOption < 1 And Not optBoolean(0).Value And Not optBoolean(1).Value And (Len(Trim$(txtAnswer.Text)) = 0) Then
            MsgBox "You haven't answered the question.", vbExclamation, "Test"
            Exit Sub
        End If
        Select Case Questions(QNo).Question_Type
          Case 1
            Questions(QNo).User_Option = UOption
          Case 2
            Questions(QNo).User_Choice = optBoolean(0).Value
          Case 3
            Questions(QNo).User_Answer = Trim$(txtAnswer.Text)
        End Select
        Questions(QNo).IsAnswered = True
        If QNo >= Test.TotalNumber Then
            GoTo Skip
        End If
      Case 2 'Quit
        If MsgBox("You haven't finished the test." & vbNewLine & "Do you really want to quit?", vbQuestion + vbYesNo, "Test") = vbYes Then
            GoTo Skip
          Else
            Exit Sub
        End If
    End Select
    UOption = 0
    QNo = QNo + 1
    ShowQuestions QNo

Exit Sub

Skip:
    ShowResult
    Unload Me

End Sub

Private Sub cmdOption_Click(Index As Integer)

  Dim i As Integer

    If cmdOption(Index).Tag = "" Then
        cmdOption(Index).Tag = "Ans"
        cmdOption(Index).BackColor = vbRed
        UOption = Index + 1
      Else
        cmdOption(Index).Tag = ""
        cmdOption(Index).BackColor = vbButtonFace
        UOption = 0
    End If
    For i = 0 To 3
        If i <> Index Then
           cmdOption(i).Tag = ""
           cmdOption(i).BackColor = vbButtonFace
        End If
    Next i

End Sub

Private Sub Form_Load()

  Dim tRs As New Recordset

    DisableX Me.hwnd
    CustomTxtBox txtAnswer, False, UpperCase
    If Test.TotalNumber = 0 Then
        MsgBox "Test options have not been set" & vbNewLine & "Contact the administrator", vbCritical, "Error starting test"
        bContinue = False
        Exit Sub
      ElseIf Test.Duration = 0 Then
        MsgBox "Test duration has been set to zero minutes" & vbNewLine & "Contact the administrator", vbCritical, "Error starting test"
        bContinue = False
        Exit Sub
      Else
        bContinue = True
    End If
    tmrCountDown.Enabled = False
    OpenRecordSet rs, "tblQuestion", , "QType, QuestionID", "SubjectID = " & Test.SubjectID

    OpenRecordSet tRs, "tblSubjects", , , "SubjectID = " & Test.SubjectID
    If tRs.RecordCount > 0 Then
        tRs.MoveFirst
        Me.Caption = "TEST ON " & tRs!SubjectName
    End If
    CloseRecordSet tRs

    Set TimePie.PictureBox = picTime
    TimePie.Value = 0
    If Test.Duration > 60 Then
        Duration = TimeSerial(Test.Duration \ 60, Test.Duration Mod 60, 0)
      Else
        Duration = TimeSerial(0, Test.Duration, 0)
    End If
    ElapsedMin = 0
    ReDim Questions(1 To Test.TotalNumber)
    QNo = 0
    Fpath = App.Path & "\NewPic.tmp"
    LoadQuestions
    QNo = 1
    ShowQuestions QNo

    tmrCountDown.Interval = 60000
    tmrCountDown.Enabled = True
    lblTime.Caption = Format$(Duration, "hh") & ":" & Format$(Duration, "nn")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseRecordSet rs
    Set TimePie = Nothing
    frmMain.picMenu.Enabled = True

End Sub

Private Sub LoadQType(Qtype As Integer, qCount As Long, Optional qStart As Long = 0)

  Dim i As Long
  Dim X As Long
  Dim Y As Long

    If qCount < 1 Then
        Exit Sub
    End If
    If qStart Then

CheckAgain:

        rs.MoveFirst
        For i = 1 To rs.RecordCount
            If Not (Y < qCount) Or Not (QNo < Test.TotalNumber) Then
                Exit For
            End If
            If rs!Qtype = Qtype Then
                X = X + 1
                If Not (X < qStart) Then
                    Y = Y + 1
                    QNo = QNo + 1
                    AddQuestion Qtype
                End If
            End If
            rs.MoveNext
        Next i
        If Y < qCount Then
            qStart = 0
            GoTo CheckAgain
        End If
      Else

        Randomize

Jump:
        If Not (Y < qCount) Or Not (QNo < Test.TotalNumber) Then
            Exit Sub
        End If
        rs.AbsolutePosition = (Rnd * (rs.RecordCount - 1)) + 1  'Get random question
        If rs!Qtype = Qtype Then
            For i = 1 To QNo
                If rs.AbsolutePosition = Questions(i).Question_No Then
                    GoTo Jump
                End If
            Next i
            Y = Y + 1
            QNo = QNo + 1
            AddQuestion Qtype
            GoTo Jump
        End If
        GoTo Jump
    End If

End Sub

Private Sub LoadQuestions()

    If Test.TotalNumber = 0 Then
        Exit Sub
    End If
    If Test.RandomDistribution Then
        LoadRandomly
        Exit Sub
    End If
    LoadQType 1, Test.MCQNo, Test.StartMCQ
    LoadQType 2, Test.TrueFalseNo, Test.StartTrueFalse
    LoadQType 3, Test.WrittenNo, Test.StartWritten

End Sub

Private Sub LoadRandomly()

  Dim i As Long

    Randomize

Jump:
    If Not (QNo < Test.TotalNumber) Then
        Exit Sub
    End If
    rs.AbsolutePosition = (Rnd * (rs.RecordCount - 1)) + 1  'Get random question
    For i = 1 To QNo
        If rs.AbsolutePosition = Questions(i).Question_No Then
            GoTo Jump
        End If
    Next i
    QNo = QNo + 1
    AddQuestion rs!Qtype
    GoTo Jump

End Sub

Private Sub picQuestion_DblClick()

    picQuestion.Zoom = 100

End Sub

Private Sub picQuestion_ZoomInClick()

    If picQuestion.Zoom < 1000 Then
        picQuestion.Zoom = picQuestion.Zoom + 10
    End If

End Sub

Private Sub picQuestion_ZoomOutClick()

    If picQuestion.Zoom > 10 Then
        picQuestion.Zoom = picQuestion.Zoom - 10
    End If

End Sub

Private Sub ShowPicture()

  'Load Picture

    On Error GoTo picError
    Set mStream = New ADODB.Stream
    mStream.Type = adTypeBinary
    mStream.Open
    mStream.Write rs!Picture.Value

    mStream.SaveToFile Fpath, adSaveCreateOverWrite
    mStream.Close
    Set mStream = Nothing

    'picQuestion.Picture = LoadPicture(Fpath)
    picQuestion.LoadImage Fpath
    Kill Fpath
    Me.Width = 12800

Exit Sub

picError:
    Me.Width = 8200

End Sub

Private Sub ShowQuestions(qCount As Long)

  Dim i As Integer

    If qCount > Test.TotalNumber Then
        Exit Sub
    End If
    ClearAll

    lblQuestNo.Caption = QNo
    lblRQuest.Caption = Test.TotalNumber - QNo
    lblTQuest.Caption = Test.TotalNumber
    If qCount = Test.TotalNumber Then
        cmdNavigate(1).Caption = "&FINISH"
      Else
        cmdNavigate(1).Caption = "&NEXT"
    End If
    rs.AbsolutePosition = Questions(qCount).Question_No
    txtQuestion.Text = rs!question
    For i = 0 To 2
        fmAns(i).Visible = False
    Next i
    fmAns(rs!Qtype - 1).Visible = True
    If rs!Qtype = 1 Then
        txtOption(0).Text = rs!Option1
        txtOption(1).Text = rs!Option2
        txtOption(2).Text = rs!Option3
        txtOption(3).Text = rs!Option4
    End If
    ShowPicture

End Sub

Private Sub ShowResult()

  Dim rsResult As New Recordset
  Dim lngResult As Long
  Dim i As Long
  Dim X As Long
  Dim Y As Long
  Dim z As Long

    X = 0
    Y = 0
    z = 0
    For i = 1 To Test.TotalNumber
        With Questions(i)
            If .IsAnswered Then
                Select Case .Question_Type
                  Case 1
                    If .Correct_Option = .User_Option Then
                        X = X + 1
                    End If
                  Case 2
                    If .Correct_Choice = .User_Choice Then
                        Y = Y + 1
                    End If
                  Case 3
                    If .Correct_Answer = .User_Answer Then
                        z = z + 1
                    End If
                End Select
            End If
        End With
    Next i
    If Test.EqualMarks Then
        lngResult = ((X + Y + z) / Test.TotalNumber) * 100
      Else
        lngResult = ((X / Test.MCQNo) * Test.MMarks)
        lngResult = lngResult + ((Y / Test.TrueFalseNo) * Test.TMarks)
        lngResult = lngResult + ((z / Test.WrittenNo) * Test.WMarks)
    End If
    OpenRecordSet rsResult, "tblScores"
    rsResult.AddNew
    rsResult!OfficialNo = User.OfficialNo
    rsResult!SubjectID = Test.SubjectID
    rsResult!Score = lngResult
    rsResult!ScoreDate = Date
    rsResult.update
    CloseRecordSet rsResult

    MsgBox "You have scored " & lngResult & " %", vbInformation, "Test Result"

End Sub

Private Sub tmrCountDown_Timer()

    If ElapsedMin = Test.Duration Then
        tmrCountDown.Enabled = False
        ShowResult
        Unload Me
        Exit Sub
      Else
        ElapsedMin = ElapsedMin + 1
    End If
    Duration = DateAdd("n", -1, Duration)
    If TimePie.Value < 25 Then
        lblTime.ForeColor = vbRed
        picTime.FillColor = vbRed
      Else
        lblTime.ForeColor = vbBlack
        picTime.FillColor = vbBlue
    End If
    TimePie.Value = ElapsedMin / Test.Duration * 100
    lblTime.Caption = Format$(Duration, "hh") & ":" & Format$(Duration, "nn")
    If ElapsedMin = Test.Duration Then
        tmrCountDown.Enabled = False
        ShowResult
        Unload Me
    End If

End Sub

