VERSION 5.00
Begin VB.Form frmPruneDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prune Data Base"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frmPruneDB.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optAllResults 
      Caption         =   "&All results."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin VB.OptionButton opt30days 
      Caption         =   "Older than 30 &days."
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
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.OptionButton opt3months 
      Caption         =   "Older than 3 &months"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
   End
   Begin VB.OptionButton opt1year 
      Caption         =   "Older than 1 &year"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdPruneRec 
      Caption         =   "&Prune Database"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPruneDB.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   7
      Top             =   480
      Width           =   7095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prune records that are:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   7095
   End
End
Attribute VB_Name = "frmPruneDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmPruneDB
' Purpose   : For pruning the results data base
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit
Private rs As New Recordset

Private Sub cmdBack_Click()

    Unload Me

End Sub

Private Sub cmdPruneRec_Click()

  Dim prunex As Long 'Number of Days to Prune
  Dim optString As String
  Dim i As Long

    If opt30days.Value Then 'If 30 days selected.
        prunex = 30 'Set Prune Length
        optString = "older than 30 days" 'Set caption
      ElseIf opt3months.Value Then  'If 1 month selected.
        prunex = 84 'Set Prune Length
        optString = "older than 3 months" 'Set Caption
      ElseIf opt1year.Value Then 'If 1 year selected.
        prunex = 365 'Set prune length
        optString = "older than 1 year" 'Set Caption.
    End If
    If optAllResults.Value Then
        If MsgBox("Are you sure you want to prune the database?" & vbNewLine & "All student results that are stored, will be lost!", vbYesNo, "Prune Database") = vbYes Then 'Confirm to user that Pruning will remove results
            OpenRecordSet rs, "tblScores"
            If rs.RecordCount > 0 Then
                GoTo FinishIt
              Else
                MsgBox "No student result stored in the data base", , "Prune data base"
                CloseRecordSet rs
                Unload Me
            End If
        End If
      ElseIf MsgBox("Are you sure you want to prune the database?" & vbNewLine & "All student results that are " & optString & ", will be lost!", vbYesNo, "Prune Database") = vbYes Then 'Confirm to user that Pruning will remove results
        OpenRecordSet rs, "tblScores", , , "ScoreDate < #" & Now - prunex & "#"
        If rs.RecordCount = 0 Then
            MsgBox "No student result is " & optString, , "Prune data base"
          Else
FinishIt:
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                rs.delete
                rs.update
                rs.MoveNext
            Next i
            MsgBox "The database has been pruned!!", vbInformation, "Prune data base" 'Inform user of success
        End If
        CloseRecordSet rs
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    DisableX Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.picMenu.Enabled = True

End Sub

