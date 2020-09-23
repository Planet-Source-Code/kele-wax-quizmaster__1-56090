VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmResult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmShow 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   10935
      Begin VB.Frame fmSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   0
         TabIndex        =   7
         Top             =   5520
         Width           =   10935
         Begin VB.Frame fmClass 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Class Analysis:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   5160
            TabIndex        =   8
            Top             =   120
            Width           =   5415
            Begin VB.Label lblComment 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "@"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   195
            End
            Begin VB.Label lblClassAvg 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1680
               TabIndex        =   10
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Class Average :"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Label lblTrend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "@"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   15
            Top             =   840
            Width           =   195
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Trend:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblAverage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   13
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Average Result:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSChart20Lib.MSChart mscResult 
         Height          =   5295
         Left            =   0
         OleObjectBlob   =   "frmResult.frx":030A
         TabIndex        =   16
         Top             =   0
         Width           =   10935
      End
      Begin MSComctlLib.ListView lvwResult 
         Height          =   5295
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SERIAL NUMBER"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TEST DATE"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SCORE (%)"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin VB.OptionButton optTable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Table"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.OptionButton optGraph 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Graph"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   8280
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   1695
   End
   Begin VB.ComboBox cboSubjects 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Summary of results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   10935
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmResult
' Purpose   : For viewing past results of a student
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset
Private ResultArr() As Long
Private iClassAvg() As Integer
Private lSubjectID() As Long
Private bDoGraph As Boolean

Private Sub cboSubjects_Change()

    If Not bDoGraph Then
        Exit Sub
    End If
    CloseRecordSet rs
    If cboSubjects.ListIndex = 0 Then
        If User.Student Then
            OpenRecordSet rs, "tblScores", , "ScoreDate, ScoreID", "OfficialNo = " & User.OfficialNo
          Else
            OpenRecordSet rs, "tblScores", , "ScoreDate, ScoreID", "OfficialNo = " & SelStudent.OfficialNo
        End If
        fmShow.Visible = True
      Else
        If User.Student Then
            OpenRecordSet rs, "tblScores", , "ScoreDate, ScoreID", "SubjectID = " & lSubjectID(cboSubjects.ListIndex) & " AND OfficialNO = " & User.OfficialNo
          Else
            OpenRecordSet rs, "tblScores", , "ScoreDate, ScoreID", "SubjectID = " & lSubjectID(cboSubjects.ListIndex) & " AND OfficialNO = " & SelStudent.OfficialNo
        End If
        If rs.RecordCount < 1 Then
            fmShow.Visible = False
            MsgBox "There are no recorded scores for " & cboSubjects.Text
            Exit Sub
          Else
            fmShow.Visible = True
        End If
    End If
    DrawGraph

End Sub

Private Sub cboSubjects_Click()

    cboSubjects_Change

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdPrint_Click()

  Dim i As Long

    On Error Resume Next
        cboSubjects.Visible = False 'Hide Subjects list
        optGraph.Visible = False 'Hide Graph option
        optTable.Visible = False 'Hide Table Option
        cmdPrint.Visible = False 'Hide Print Button
        cmdClose.Visible = False 'Hide Close Button
        'PrintForm 'Print the form
        If Not mscResult.Visible Then
            If MsgBox("Do you want to print all the results from the begining to the end?", vbQuestion + vbYesNo, "Print results") = vbYes Then
                lvwResult.ListItems(1).EnsureVisible
                PrintForm
                If rs.RecordCount > 18 Then
                    For i = 1 To rs.RecordCount \ 18
                        Scroll
                        PrintForm
                    Next i
                End If
              Else
                PrintForm
            End If
          Else
            PrintForm
        End If

        cboSubjects.Visible = True 'After sending to printer Show
        optGraph.Visible = True 'After sending to printer Show
        optTable.Visible = True 'After sending to printer Show
        cmdClose.Visible = True 'After sending to printer Show
        cmdPrint.Visible = True 'After sending to printer Show
    On Error GoTo 0

End Sub

Private Sub DrawGraph()

  Dim i As Long
  Dim tScore As Long
  Dim avgScore As Integer
  Dim fScore As Integer
  Dim LScore As Integer
  Dim mScore As Integer
  Dim fTrend As Integer
  Dim LTrend As Integer
  Dim Trend As Integer
  Dim cLvw As ListItem

    If User.Student Then
        lblCaption.Caption = "RESULT SUMMARY FOR " & User.Rank & " " & User.Initials & " " & User.Surname & " ON " & cboSubjects.Text
      Else
        lblCaption.Caption = "RESULT SUMMARY FOR " & SelStudent.Rank & " " & SelStudent.Initials & " " & SelStudent.Surname & " ON " & cboSubjects.Text
    End If

    mscResult.ColumnCount = 1
    mscResult.RowCount = rs.RecordCount
    lvwResult.ListItems.Clear
    ReDim ResultArr(rs.RecordCount)
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        Set cLvw = lvwResult.ListItems.Add
        mscResult.Column = 1
        mscResult.Row = i
        cLvw.Text = i
        mscResult.RowLabel = CStr(rs!ScoreDate)
        cLvw.SubItems(1) = CStr(Format$(rs!ScoreDate, "dddd, dd mmmm yyyy"))
        mscResult.Data = rs!Score
        cLvw.SubItems(2) = rs!Score
        ResultArr(i) = rs!Score
        tScore = tScore + rs!Score
        rs.MoveNext
    Next i

    If rs.RecordCount Mod 18 > 0 Then
        For i = (rs.RecordCount + 1) To (rs.RecordCount + (18 - (rs.RecordCount Mod 18)))
            Set cLvw = lvwResult.ListItems.Add
            cLvw.Text = "  "
        Next i
    End If
    avgScore = tScore / rs.RecordCount
    lblAverage.Caption = avgScore
    lblClassAvg.Caption = iClassAvg(cboSubjects.ListIndex)

    If iClassAvg(cboSubjects.ListIndex) < avgScore Then 'If Student Average more than population average then...
        lblComment.Caption = "Student is above average in class." 'Comment.
      ElseIf iClassAvg(cboSubjects.ListIndex) = avgScore Then 'If Student Average the same as population average then...
        lblComment.Caption = "Student is average in class." 'Comment.
      Else 'If Student Average less than poppulation average then...
        lblComment.Caption = "Student is below average in class." 'Comment.
    End If

    fScore = ResultArr(1)
    mScore = (ResultArr(UBound(ResultArr) / 2)) * 2
    LScore = ResultArr(UBound(ResultArr))

    fTrend = mScore - fScore
    LTrend = LScore - mScore
    Trend = fTrend + LTrend

    If Trend > 0 Then
        lblTrend.Caption = "Improving Results"
      ElseIf Trend < 0 Then
        lblTrend.Caption = "Deteriorating Results"
      Else
        lblTrend.Caption = "Consistent Results"
    End If

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd
    If User.Student Then
        OpenRecordSet rs, "tblScores", , "ScoreDate, ScoreID", "OfficialNo = " & User.OfficialNo
      Else
        OpenRecordSet rs, "tblScores", , "ScoreDate, ScoreID", "OfficialNo = " & SelStudent.OfficialNo
    End If
    If rs.RecordCount < 1 Then
        CloseRecordSet rs
        MsgBox "There are no record of any test score.", vbInformation, "Result"
        bContinue = False
      Else
        bContinue = True
        LoadSubjects
        GetClassAvg
        cboSubjects.ListIndex = 0
        DrawGraph
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseRecordSet rs
    Erase ResultArr
    Erase iClassAvg
    Erase lSubjectID
    If User.Student Then
        frmMain.picMenu.Enabled = True
      Else
        frmUserInformation.Show
    End If

End Sub

Private Sub GetClassAvg()

  Dim rsScores As New Recordset
  Dim i As Long
  Dim X As Long
  Dim tScore As Long
  Dim tStudent As Long

    If UBound(lSubjectID) > 0 Then
        ReDim iClassAvg(UBound(lSubjectID))
        iClassAvg(0) = 0
        OpenRecordSet rsScores, "tblScores"
        If rsScores.RecordCount > 0 Then
            rsScores.MoveFirst
            For i = 1 To rsScores.RecordCount
                tScore = tScore + rsScores!Score
                rsScores.MoveNext
            Next i
            iClassAvg(0) = tScore / rsScores.RecordCount
          Else
            MsgBox "No results are stored in the database", vbInformation, "Error viewing result"
            Unload Me
        End If
        CloseRecordSet rsScores

        For i = 1 To UBound(lSubjectID)
            iClassAvg(i) = 0
            tStudent = 0
            tScore = 0
            OpenRecordSet rsScores, "tblScores", , , "SubjectID = " & lSubjectID(i)
            If rsScores.RecordCount > 0 Then
                rsScores.MoveFirst
                For X = 1 To rsScores.RecordCount
                    tScore = tScore + rsScores!Score
                    tStudent = tStudent + 1
                    rsScores.MoveNext
                Next X
                iClassAvg(i) = tScore / tStudent
            End If
            CloseRecordSet rsScores
        Next i
    End If

End Sub

Private Sub LoadSubjects()

  Dim rsSubject As New Recordset
  Dim i As Long

    bDoGraph = False
    cboSubjects.Clear
    OpenRecordSet rsSubject, "tblSubjects"
    If rsSubject.RecordCount > 0 Then
        rsSubject.MoveFirst
        ReDim lSubjectID(rsSubject.RecordCount)
        cboSubjects.AddItem "All Test", 0
        lSubjectID(0) = 0
        For i = 1 To rsSubject.RecordCount
            cboSubjects.AddItem rsSubject!SubjectName, i
            lSubjectID(i) = rsSubject!SubjectID
            rsSubject.MoveNext
        Next i
        CloseRecordSet rsSubject
      Else
        MsgBox "There are no register subjects so, you cant see any results", vbInformation, "Error showing results"
        Unload Me
    End If
    bDoGraph = True

End Sub

Private Sub optGraph_Click()

    mscResult.Visible = optGraph.Value
    lvwResult.Visible = Not optGraph.Value

End Sub

Private Sub optTable_Click()

    lvwResult.Visible = optTable.Value
    mscResult.Visible = Not optTable.Value

End Sub

