Attribute VB_Name = "modGlobal"
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.

Option Explicit

' Declaration for Stay on Top sub
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Const meScroll = &H1000 + 20

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Enum ChrCase ' For Customization of textbox(Character Case)
    AnyCase = 0
    UpperCase = 1
    LowerCase = 2
End Enum

Private Type TestInfo
    SubjectID As Long
    TotalNumber As Long
    Duration As Long
    RandomDistribution As Boolean
    MCQNo As Long
    TrueFalseNo As Long
    WrittenNo As Long
    StartMCQ As Long
    StartTrueFalse As Long
    StartWritten As Long
    EqualMarks As Boolean
    MMarks As Long
    TMarks As Long
    WMarks As Long
End Type

Public Test As TestInfo

Private Type UserInfo
    Rank As String
    Initials As String
    Surname As String
    OfficialNo As Long
    Student As Boolean
    Password As String
End Type

Public User As UserInfo
Public SelStudent As UserInfo  'Selected Student Information(In Administrator Mode)

Private Type Quest
    Question_No As Integer
    Question_Type As Integer
    Correct_Option As Integer
    Correct_Choice As Boolean
    Correct_Answer As String
    IsAnswered As Boolean
    User_Option As Integer
    User_Answer As String
    User_Choice As Boolean
End Type

Public Questions() As Quest

'*********** INTEGER VARIABLE *************
Public QuestAsked As Long 'To keep tract of number of questions asked
Public Incre As Long 'Counter variable for random question generation
Public Score As Long 'To keep tract of marks the student is scorring

'********* BOOLEAN VARIABLE ***********'

Public bContinue As Boolean
Public AddQuestion As Boolean

Public Function CanContinue(frm As Form) As Boolean

    CanContinue = bContinue
    If bContinue Then
        frm.Show
      Else
        Unload frm
    End If

End Function

Public Sub CloseRecordSet(rs As Recordset)

    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Sub

Sub CustomTxtBox(TextBox As TextBox, Optional Numeric As Boolean = True, Optional CharacterCase As ChrCase = AnyCase)

  Dim style As Long
  Const GWL_STYLE = (-16)
  Const ES_NUMBER = &H2000
  Const ES_UPPERCASE = &H8&
  Const ES_LOWERCASE = &H10&

    ' get current style
    style = GetWindowLong(TextBox.hwnd, GWL_STYLE)
    If Numeric Then
        'Allow only numeric
        style = style Or ES_NUMBER
      Else
        ' Allow any character
        style = style And Not ES_NUMBER
    End If

    Select Case CharacterCase
      Case 0
        ' restore default style
        style = style And Not (ES_UPPERCASE Or ES_LOWERCASE)
      Case 1
        ' convert to uppercase
        style = style Or ES_UPPERCASE
      Case 2
        ' convert to lowercase
        style = style Or ES_LOWERCASE
    End Select
    ' enforce new style
    SetWindowLong TextBox.hwnd, GWL_STYLE, style

End Sub

Public Sub DisableX(hwnd As Long, Optional fEnable As Boolean = False)

  ' Disable X in upper right corner of the form

  Dim lngMenu As Long

    lngMenu = GetSystemMenu(hwnd, fEnable)
    If Not fEnable Then
        DeleteMenu lngMenu, 6, MF_BYPOSITION
    End If
    DrawMenuBar hwnd

End Sub

Public Sub GetTestOptions()

  Dim rs As New Recordset

    With Test
        .Duration = 0
        .EqualMarks = True
        .MCQNo = 0
        .MMarks = 0
        .RandomDistribution = True
        .StartMCQ = 0
        .StartTrueFalse = 0
        .StartWritten = 0
        .SubjectID = 0
        .TMarks = 0
        .TotalNumber = 0
        .TrueFalseNo = 0
        .WMarks = 0
        .WrittenNo = 0
    End With

    OpenRecordSet rs, "tblAdmin"
    On Error Resume Next
        rs.MoveFirst
        With Test
            .Duration = rs!Duration
            .MCQNo = rs!MCQNo
            .RandomDistribution = rs!Ran_Dist
            .StartMCQ = rs!StartMCQ
            .StartTrueFalse = rs!StartTrueFalse
            .StartWritten = rs!StartWritten
            .SubjectID = rs!SubjectID
            .TotalNumber = rs!TotalNumber
            .TrueFalseNo = rs!TrueFalseNo
            .WrittenNo = rs!WrittenNo
            .EqualMarks = rs!Equal
            .MMarks = rs!MCQ
            .TMarks = rs!TrueFalse
            .WMarks = rs!Written
        End With
    On Error GoTo 0
    CloseRecordSet rs

End Sub

Public Sub OnTop(TheForm As Form, Optional bUp As Boolean = True)

  ' Put window on top

    If bUp Then

        SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
      Else

        SetWindowPos TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

End Sub

Public Sub OpenRecordSet(tRs As Recordset, Table As String, Optional Scope As String = "*", Optional OrderBy As String = "", Optional Filter As String = "") 'As Recordset

  Dim Path As String
  Dim ConnectionString As String

    If OrderBy <> "" Then
        OrderBy = "Order By " & OrderBy
    End If

    If Filter <> "" Then
        Filter = "Where " & Filter
    End If

    Path = App.Path & "\QuizMaster.0_0"
    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path & ";Jet OLEDB:Database Password=kelewax2165asm1299"

    Set tRs = New Recordset
    tRs.CursorLocation = adUseClient
    tRs.ActiveConnection = ConnectionString
    tRs.CursorType = adOpenDynamic
    tRs.LockType = adLockOptimistic
    tRs.Source = "Select " & Scope & " from " & Table & " " & Filter & " " & OrderBy
    tRs.Open

End Sub

Public Sub Scroll()

  Dim ScrollY As Long

    ScrollY = 4620 / Screen.TwipsPerPixelY

    SendMessage frmResult.lvwResult.hwnd, meScroll, 0, ByVal ScrollY

End Sub

