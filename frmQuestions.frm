VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQuestions 
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   Icon            =   "frmQuestions.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12975
   Begin prjQuizMaster.ZoomPicCtl picQuestion 
      Height          =   6255
      Left            =   8040
      TabIndex        =   42
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   11033
      BackColor       =   -2147483633
      AllowZoomIn     =   -1  'True
      AllowZoomOut    =   -1  'True
      UseQuickBar     =   -1  'True
   End
   Begin VB.Frame fmAns 
      Height          =   3735
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   2880
      Width           =   7695
      Begin VB.OptionButton optMCQ 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   495
      End
      Begin VB.OptionButton optMCQ 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   495
      End
      Begin VB.OptionButton optMCQ 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.OptionButton optMCQ 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   0
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   2
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2160
         Width           =   6495
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   3
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2880
         Width           =   6495
      End
      Begin VB.TextBox txtOption 
         Height          =   600
         Index           =   1
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   6495
      End
      Begin VB.Label Label2 
         Caption         =   "OPTIONS"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cdcPhoto 
      Left            =   7200
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Picture File"
      Filter          =   "All Files(*.*)|*.*|All Image Files|*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.cur"
      FilterIndex     =   1
      MaxFileSize     =   100
   End
   Begin VB.Frame fmQuestion 
      Height          =   2535
      Left            =   240
      TabIndex        =   29
      Top             =   240
      Width           =   7695
      Begin VB.TextBox txtQuestion 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Label Label6 
         Caption         =   "Question No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Remaining Question :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   39
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Total Questions :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblQuestNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblTQuest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6840
         TabIndex        =   36
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblRQuest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4560
         TabIndex        =   35
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Timer tmrCountDown 
      Left            =   9000
      Top             =   4440
   End
   Begin VB.Frame fmAns 
      Height          =   855
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   2880
      Width           =   7695
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
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
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
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fmAns 
      Height          =   1815
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   2880
      Width           =   7695
      Begin VB.TextBox txtAnswer 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label Label3 
         Caption         =   "ANSWER"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fmNavigate 
      Height          =   1455
      Left            =   240
      TabIndex        =   27
      Top             =   6600
      Width           =   7695
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&Jump Forward"
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
         Index           =   2
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Jump Bac&k"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdNavigate 
         Cancel          =   -1  'True
         Caption         =   "&Close"
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
         Index           =   8
         Left            =   6240
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&Update"
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
         Index           =   6
         Left            =   3360
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&Delete"
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
         Index           =   4
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&Edit"
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
         Index           =   5
         Left            =   1920
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "P&revious"
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
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "&Next"
         Default         =   -1  'True
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
         Index           =   3
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Ne&w"
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
         Index           =   7
         Left            =   4800
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblJumpBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3840
         TabIndex        =   41
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame fmNewQ 
      Height          =   1455
      Left            =   240
      TabIndex        =   34
      Top             =   6600
      Width           =   7695
      Begin VB.OptionButton optQType 
         Caption         =   "Written"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optQType 
         Caption         =   "True or False"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optQType 
         Caption         =   "Multiple Choice"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   24
         Top             =   960
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton cmdNQuest 
         Caption         =   "&Back"
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
         Index           =   2
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdNQuest 
         Caption         =   "&Picture"
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
         Index           =   0
         Left            =   720
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdNQuest 
         Caption         =   "&Save"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmQuestions
' Purpose   : For viewing, adding, editing and deleting questions on a particular subject
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Public SubjectID As Long
Public rs As New Recordset
Private mStream As New ADODB.Stream
Private Fpath As String

Private Sub ClearAll()

  Dim i As Long

    lblQuestNo.Caption = ""
    lblRQuest.Caption = ""
    lblTQuest.Caption = ""
    txtQuestion.Text = ""
    For i = 0 To 3
        txtOption(i).Text = ""
        optMCQ(i).Value = False
    Next i
    txtAnswer.Text = ""
    optBoolean(0).Value = False
    optBoolean(1).Value = False
    'picQuestion.Picture = LoadPicture
    picQuestion.UnloadImage
    Me.Width = 8200
    fmAns(0).Visible = True
    fmAns(1).Visible = False
    fmAns(2).Visible = False

End Sub

Private Sub cmdNavigate_Click(Index As Integer)

  Dim pp As Long
  Dim j As Integer
  Dim i As Long
  Dim tRs As New Recordset

    If Not fmNavigate.Visible Then
        Exit Sub
    End If

    j = Val(lblJumpBy)

    Select Case Index
      Case 0, 1, 2, 3, 4, 5, 6
        If rs.RecordCount < 1 Then
            ClearAll
            MsgBox "There are no questions in the database", vbExclamation, "Error message"
            Exit Sub
        End If
        If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Then
            If cmdNavigate(5).Caption = "Cance&l" Then
                MsgBox "Confirm your changes first", vbCritical, "Error message"
                Exit Sub
            End If
        End If

        Select Case Index
          Case 0 'Previous
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveLast
            End If

          Case 1    'Jump Back
            If rs.AbsolutePosition > j Then
                rs.AbsolutePosition = rs.AbsolutePosition - j
              Else
                rs.MoveFirst
            End If
          Case 2    'Jump Forward
            If rs.AbsolutePosition + j < rs.RecordCount Then
                rs.AbsolutePosition = rs.AbsolutePosition + j
              Else
                rs.MoveLast
            End If
          Case 3    'Next
            rs.MoveNext
            If rs.EOF Then
                rs.MoveFirst
            End If
          Case 4    'Delete and Picture

            If cmdNavigate(4).Caption = "&Delete" Then
                If MsgBox("Are you sure you want to delete question number " & rs.AbsolutePosition & "?" & vbNewLine & "You would have to reset test options", vbYesNo, "Delete question") = vbYes Then
                    pp = rs.AbsolutePosition
                    rs.delete
                    rs.update
                    CloseRecordSet rs
                    OpenRecordSet rs, "tblQuestion", , "QType, QuestionID", "SubjectID = " & SubjectID
                    If rs.RecordCount = 0 Then
                        ShowQuestion True
                        Exit Sub
                      ElseIf rs.RecordCount = 1 Then
                        rs.MoveFirst
                      ElseIf pp < rs.RecordCount Then
                        rs.AbsolutePosition = pp
                      ElseIf pp > rs.RecordCount Then
                        rs.MoveLast
                    End If

                    OpenRecordSet tRs, "tblAdmin"
                    If tRs.RecordCount > 0 Then
                        tRs.MoveFirst
                        tRs.delete
                        tRs.update
                    End If
                    CloseRecordSet tRs
                  Else
                    Exit Sub
                End If
              ElseIf cmdNavigate(4).Caption = "Pic&ture" Then

                If cmdNavigate(5).Caption = "&Edit" Then
                    Exit Sub
                End If

                On Error GoTo picError  'In case the photopath cant be displayed
                cdcPhoto.ShowOpen
                If Len(cdcPhoto.FileName) <> 0 Then
                    FileCopy cdcPhoto.FileName, Fpath 'save picture in a temp file
                End If

                'picQuestion.Picture = LoadPicture(Fpath) 'load picture from temp file
                picQuestion.LoadImage Fpath
                Me.Width = 12800
                Exit Sub
picError:
                If (Err.Number <> 0) And (Err.Number <> 32755) Then
                    MsgBox "Error number " & Err.Number & ":- " & Err.Description, , "Error Loading Picture"
                    Me.Width = 8200
                End If
                Exit Sub

            End If
          Case 5    'Edit or Cancel
            If cmdNavigate(5).Caption = "&Edit" Then
                MsgBox "Make your changes and click the Update button"
                cmdNavigate(5).Caption = "Cance&l"
                cmdNavigate(4).Caption = "Pic&ture"
                cmdNavigate(8).Enabled = False
                UnlockAll
                Exit Sub
              ElseIf cmdNavigate(5).Caption = "Cance&l" Then
                cmdNavigate(5).Caption = "&Edit"
                cmdNavigate(4).Caption = "&Delete"
                cmdNavigate(8).Enabled = True
                LockAll
            End If

          Case 6    'Update
            If cmdNavigate(5).Caption = "&Edit" Then
                Exit Sub
            End If
            If txtQuestion.Text = "" Then
                MsgBox "You haven't entered a question.", , "Error Editing Question."
                txtQuestion.SetFocus
                Exit Sub
            End If
            Select Case rs!Qtype
              Case 1
                For i = 0 To 3
                    If Trim$(txtOption(i).Text) = "" Then
                        MsgBox "You haven't entered option " & i + 1, , "Error Editing Queusion."
                        txtOption(i).SetFocus
                        Exit Sub
                    End If
                Next i
                If Not (optMCQ(0).Value Or optMCQ(1).Value Or optMCQ(2).Value Or optMCQ(3).Value) Then
                    MsgBox "You haven't selected the correct option.", , "Error editing question"
                    Exit Sub
                End If
              Case 2
                If Not optBoolean(0).Value And Not optBoolean(1).Value Then
                    MsgBox "You haven't selected the correct option.", , "Error editing question"
                    Exit Sub
                End If
              Case 3
                If Trim$(txtAnswer.Text) = "" Then
                    MsgBox "You haven't entered the answer.", , "Error editing question"
                    txtAnswer.SetFocus
                    Exit Sub
                End If
              Case Else
                MsgBox "Sorry the operation could not be completed. Quit", , "Error editing question"
                Exit Sub

            End Select
            rs!question = Trim$(txtQuestion.Text)
            rs!Option1 = Trim$(txtOption(0).Text)
            rs!Option2 = Trim$(txtOption(1).Text)
            rs!Option3 = Trim$(txtOption(2).Text)
            rs!Option4 = Trim$(txtOption(3).Text)
            For i = 0 To 3
                If optMCQ(i).Value Then
                    rs!CorrectOption = i + 1
                    Exit For
                End If
            Next i
            rs!Boolean = optBoolean(0).Value
            rs!Answer = Trim$(txtAnswer.Text)
            rs.update
            On Error GoTo picErr 'Resume Next
            'Save the picture as binary data from the temp file
            Set mStream = New ADODB.Stream
            mStream.Type = adTypeBinary
            mStream.Open
            mStream.LoadFromFile Fpath
            rs!Picture.Value = mStream.Read
            mStream.Close
            rs.update
            Set mStream = Nothing
            Kill Fpath 'delete the temp file
picErr:
            '  On Error GoTo 0 'clear error handling
            cmdNavigate(5).Caption = "&Edit"
            cmdNavigate(4).Caption = "&Delete"
            cmdNavigate(8).Enabled = True

        End Select
        ShowQuestion

      Case 7    'New
        For i = 0 To 2
            optQType(i).Value = fmAns(i).Visible
        Next i
        optQType(0).Value = True
        fmNewQ.Visible = True
        fmNavigate.Visible = False
        ClearAll
        UnlockAll
        MsgBox "Type in your questions with their answers and click the 'Save' button", vbInformation, "New Questions"

      Case 8    'Close
        frmSubjects.Show
        Unload Me
    End Select

End Sub

Private Sub cmdNQuest_Click(Index As Integer)

  Dim i As Long

    If Not fmNewQ.Visible Then
        Exit Sub
    End If
    Select Case Index
      Case 0

        On Error GoTo picError  'In case the photopath cant be displayed
        cdcPhoto.ShowOpen
        If Len(cdcPhoto.FileName) <> 0 Then
            FileCopy cdcPhoto.FileName, Fpath 'save picture in a temp file
        End If

        'picQuestion.Picture = LoadPicture(Fpath) 'load picture from temp file
        picQuestion.LoadImage Fpath
        Me.Width = 12800
        Exit Sub
picError:
        Me.Width = 8200
        If (Err.Number <> 0) And (Err.Number <> 32755) Then
            MsgBox "Error number " & Err.Number & ":- " & Err.Description, , "Error Loading Picture"
        End If

      Case 1
        If txtQuestion.Text = "" Then
            MsgBox "You haven't entered a question.", , "Error Writting Question."
            txtQuestion.SetFocus
            Exit Sub
        End If
        If optQType(0).Value Then
            For i = 0 To 3
                If Trim$(txtOption(i).Text) = "" Then
                    MsgBox "You haven't entered option " & i + 1, , "Error Writting Queusion."
                    txtOption(i).SetFocus
                    Exit Sub
                End If
            Next i
            If Not (optMCQ(0).Value Or optMCQ(1).Value Or optMCQ(2).Value Or optMCQ(3).Value) Then
                MsgBox "You haven't selected the correct option.", , "Error Writting Question"
                Exit Sub
            End If
          ElseIf optQType(1).Value Then

            If Not optBoolean(0).Value And Not optBoolean(1).Value Then
                MsgBox "You haven't selected the correct option.", , "Error Writting Question"
                Exit Sub
            End If
          ElseIf optQType(2).Value Then

            If Trim$(txtAnswer.Text) = "" Then
                MsgBox "You haven't entered the answer.", , "Error Writting Question"
                txtAnswer.SetFocus
                Exit Sub
            End If
          Else
            MsgBox "Sorry the operation could not be completed. Quit", , "Error Writting Question"
            Exit Sub

        End If
        rs.AddNew
        rs!SubjectID = SubjectID
        rs!question = Trim$(txtQuestion.Text)
        If optQType(0).Value Then
            rs!Option1 = Trim$(txtOption(0).Text)
            rs!Option2 = Trim$(txtOption(1).Text)
            rs!Option3 = Trim$(txtOption(2).Text)
            rs!Option4 = Trim$(txtOption(3).Text)
            For i = 0 To 3
                If optMCQ(i).Value Then
                    rs!CorrectOption = i + 1
                    Exit For
                End If
            Next i
            rs!Qtype = 1
          ElseIf optQType(1).Value Then
            rs!Boolean = optBoolean(0).Value
            rs!Qtype = 2
          ElseIf optQType(2).Value Then
            rs!Answer = Trim$(txtAnswer.Text)
            rs!Qtype = 3
        End If

        'Save the picture as binary data from the temp file
        On Error Resume Next
            Set mStream = New ADODB.Stream
            mStream.Type = adTypeBinary
            mStream.Open
            mStream.LoadFromFile Fpath
            rs!Picture.Value = mStream.Read
            mStream.Close
            Set mStream = Nothing
            Kill Fpath 'delete the temp file
        On Error GoTo 0 'clear error handling
        rs.update
        ClearAll
        optQType(0).Value = True
      Case 2
        LockAll
        fmNewQ.Visible = False
        fmNavigate.Visible = True
        ShowQuestion True
    End Select

End Sub

Private Sub Form_Load()

  Dim tRs As New Recordset

    DisableX Me.hwnd
    CustomTxtBox txtAnswer, False, UpperCase
    Me.Height = 8800
    Fpath = App.Path & "\NewPic.tmp"
    OpenRecordSet rs, "tblQuestion", , "QType, QuestionID", "SubjectID = " & SubjectID

    OpenRecordSet tRs, "tblSubjects", , , "SubjectID = " & SubjectID
    If tRs.RecordCount > 0 Then
        tRs.MoveFirst
        Me.Caption = "SUBJECT :-  " & tRs!SubjectName
    End If
    CloseRecordSet tRs

    ShowQuestion True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseRecordSet rs
    frmSubjects.Show

End Sub

Private Sub lblJumpBy_DblClick()

  Dim j As String

    j = InputBox("Enter a number to jump by.", "Questions")
    If Val(j) < 1 Then
        lblJumpBy.Caption = "1"
      Else
        lblJumpBy.Caption = CStr(Val(j))
    End If

End Sub

Private Sub LockAll()

    fmQuestion.Enabled = False
    fmAns(0).Enabled = False
    fmAns(1).Enabled = False
    fmAns(2).Enabled = False
    'picQuestion.Enabled = False

End Sub

Private Sub optQType_Click(Index As Integer)

  Dim i As Integer

    For i = 0 To 2
        fmAns(i).Visible = optQType(i).Value
    Next i
    Select Case Index
      Case 0
        optBoolean(0).Value = False
        optBoolean(1).Value = False
        txtAnswer.Text = ""
      Case 1
        For i = 0 To 3
            optMCQ(i).Value = False
            txtOption(i).Text = ""
        Next i
        txtAnswer.Text = ""
      Case 2
        For i = 0 To 3
            optMCQ(i).Value = False
            txtOption(i).Text = ""
        Next i
        optBoolean(0).Value = False
        optBoolean(1).Value = False
    End Select

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

Private Sub ShowQuestion(Optional First As Boolean = False)

    ClearAll
    LockAll
    If First Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
          Else
            Exit Sub
        End If
    End If

    SubjectID = rs!SubjectID
    lblQuestNo.Caption = rs.AbsolutePosition
    lblRQuest.Caption = rs.RecordCount - rs.AbsolutePosition
    lblTQuest.Caption = rs.RecordCount

    txtQuestion.Text = rs!question

    If rs!Qtype = 1 Then
        txtOption(0).Text = rs!Option1
        txtOption(1).Text = rs!Option2
        txtOption(2).Text = rs!Option3
        txtOption(3).Text = rs!Option4
        optMCQ(rs!CorrectOption - 1).Value = True
      ElseIf rs!Qtype = 2 Then
        optBoolean(0).Value = rs!Boolean
        optBoolean(1).Value = Not rs!Boolean
      ElseIf rs!Qtype = 3 Then
        txtAnswer.Text = rs!Answer
    End If

    fmAns(0).Visible = False
    fmAns(1).Visible = False
    fmAns(2).Visible = False
    fmAns(rs!Qtype - 1).Visible = True

    'Load Picture
    On Error GoTo picError
    Set mStream = New ADODB.Stream
    mStream.Type = adTypeBinary
    mStream.Open
    mStream.Write rs!Picture.Value

    mStream.SaveToFile Fpath, adSaveCreateOverWrite
    mStream.Close
    Set mStream = Nothing

    picQuestion.LoadImage Fpath
    Kill Fpath
    Me.Width = 12800

Exit Sub

picError:
    Me.Width = 8200

End Sub

Private Sub UnlockAll()

    fmQuestion.Enabled = True
    fmAns(0).Enabled = True
    fmAns(1).Enabled = True
    fmAns(2).Enabled = True

End Sub

