VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "QUIZ MASTER"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMenu 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7080
      ScaleWidth      =   1590
      TabIndex        =   0
      Top             =   0
      Width           =   1590
      Begin VB.CommandButton cmdMenu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   100
         TabIndex        =   1
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain
' Purpose   : The main display form, for navigating to other parts of the programme
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdMenu_Click(Index As Integer)

    If User.Student Then
        Select Case Index
          Case 0
            GetTestOptions
            Load frmTest
            If Not (CanContinue(frmTest)) Then
                Exit Sub
            End If
          Case 1
            Load frmResult
            If Not (CanContinue(frmResult)) Then
                Exit Sub
            End If
          Case 2
            Load frmChangePwd
            frmChangePwd.Show
          Case 3
            Load frmLogin
            frmLogin.Show
            Unload Me
            Exit Sub
          Case 4
            Unload Me
            Exit Sub
          Case 5
            OnTop Me, False
            Load frmHelp
            frmHelp.Show
        End Select
      Else
        Select Case Index
          Case 0
            Load frmUserInformation
            frmUserInformation.Show
          Case 1
            Load frmSubjects
            frmSubjects.Show
          Case 2
            Load frmSelectQuestion
            frmSelectQuestion.Show
          Case 3
            Load frmPruneDB
            frmPruneDB.Show
          Case 4
            Load frmChangePwd
            frmChangePwd.Show
          Case 5
            Load frmLogin
            frmLogin.Show
            Unload Me
            Exit Sub
          Case 6
            Unload Me
            Exit Sub
          Case 7
            OnTop Me, False
            Load frmHelp
            frmHelp.Show
        End Select
    End If

    picMenu.Enabled = False

End Sub

Private Sub LoadMenu(Student As Boolean)

  Dim i As Integer

    cmdMenu(0).Top = 200
    If Student Then
        For i = 1 To 5
            Load cmdMenu(i)
            cmdMenu(i).Visible = True
            cmdMenu(i).Top = cmdMenu(i - 1).Top + cmdMenu(i - 1).Height + 200
        Next i

        cmdMenu(0).Caption = "Start" & vbNewLine & "Qui&z"
        cmdMenu(1).Caption = "View" & vbNewLine & "&Result"
        cmdMenu(2).Caption = "Change" & vbNewLine & "&Password"
        cmdMenu(3).Caption = "&Log Out"
        cmdMenu(4).Caption = "&Quit"
        cmdMenu(4).Cancel = True
        cmdMenu(5).Caption = "&Help"
      Else
        For i = 1 To 7
            Load cmdMenu(i)
            cmdMenu(i).Visible = True
            cmdMenu(i).Top = cmdMenu(i - 1).Top + cmdMenu(i - 1).Height + 200
        Next i

        cmdMenu(0).Caption = "Manage &Users" & vbCrLf & "and" & vbCrLf & "View Results"
        cmdMenu(1).Caption = "Manage Qui&z" & vbNewLine & "Library"
        cmdMenu(2).Caption = "Select Quiz" & vbNewLine & "&Options"
        cmdMenu(3).Caption = "Prune Results" & vbNewLine & "&Database"
        cmdMenu(4).Caption = "Change" & vbNewLine & "&Password"
        cmdMenu(5).Caption = "&Log Out"
        cmdMenu(6).Caption = "&Quit"
        cmdMenu(6).Cancel = True
        cmdMenu(7).Caption = "&Help"

    End If

End Sub

Private Sub MDIForm_Load()

    OnTop Me
    DisableX Me.hwnd
    Me.WindowState = 2
    LoadMenu User.Student

End Sub

