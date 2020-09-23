Attribute VB_Name = "modCreateDataBase"
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.

Option Explicit

Private nClsDb As clsCreateDataBase

Public Sub CreateNewDataBase()

  Dim myPath As String

    myPath = App.Path & "\QuizMaster.0_0"
    Set nClsDb = New clsCreateDataBase
    nClsDb.DBConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & myPath & ";Jet OLEDB:Database Password=kelewax2165asm1299;Jet OLEDB:Engine Type=5;"
    If Not nClsDb.CreateDB(myPath, False) Then Exit Sub
    nClsDb.CreateTable "tblAdmin"
    nClsDb.CreateTable "tblScores"
    nClsDb.CreateTable "tblSubjects"
    nClsDb.CreateTable "tblQuestion"
    nClsDb.CreateTable "tblUsers"

    nClsDb.CreateColumn "tblAdmin", "SubjectID", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "TotalNumber", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "Duration", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "MCQNo", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "TrueFalseNo", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "WrittenNo", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "StartMCQ", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "StartTrueFalse", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "StartWritten", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "MCQ", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "TrueFalse", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "Written", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblAdmin", "Ran_Dist", adBoolean
    nClsDb.CreateColumn "tblAdmin", "Equal", adBoolean

    nClsDb.CreateColumn "tblQuestion", "QuestionID", adInteger, 0, 1, 0
    nClsDb.CreateColumn "tblQuestion", "SubjectID", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "CorrectOption", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "QType", adInteger, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "Question", adLongVarWChar, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "Answer", adLongVarWChar, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "Boolean", adBoolean
    nClsDb.CreateColumn "tblQuestion", "Picture", adLongVarBinary, , , 1
    nClsDb.CreateColumn "tblQuestion", "Option1", adVarWChar, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "Option2", adVarWChar, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "Option3", adVarWChar, 0, 0, 1
    nClsDb.CreateColumn "tblQuestion", "Option4", adVarWChar, 0, 0, 1

    nClsDb.CreateColumn "tblScores", "ScoreID", adInteger, 0, 1
    nClsDb.CreateColumn "tblScores", "SubjectID", adInteger, 0, , 1
    nClsDb.CreateColumn "tblScores", "OfficialNo", adInteger, 0, , 1
    nClsDb.CreateColumn "tblScores", "Score", adInteger, 0, , 1
    nClsDb.CreateColumn "tblScores", "ScoreDate", adDate, , , 1

    nClsDb.CreateColumn "tblSubjects", "SubjectID", adInteger, 0, 1
    nClsDb.CreateColumn "tblSubjects", "SubjectName", adVarWChar, , , 1

    nClsDb.CreateColumn "tblUsers", "OfficialNo", adInteger, 0, , 1
    nClsDb.CreateColumn "tblUsers", "Rank", adVarWChar, , , 1
    nClsDb.CreateColumn "tblUsers", "Initials", adVarWChar, , , 1
    nClsDb.CreateColumn "tblUsers", "Surname", adVarWChar, , , 1
    nClsDb.CreateColumn "tblUsers", "Password", adVarWChar, , , 1
    nClsDb.CreateColumn "tblUsers", "Student", adBoolean
    nClsDb.CreateColumn "tblUsers", "Active", adBoolean

    nClsDb.CreatePrimaryKey "tblSubjects", "SubjectID"
    nClsDb.CreatePrimaryKey "tblQuestion", "QuestionID"
    nClsDb.CreatePrimaryKey "tblScores", "ScoreID"
    nClsDb.CreatePrimaryKey "tblUsers", "OfficialNo"

    nClsDb.CreateIndex "SubjectID", "tblQuestion", "SubjectID", False, adIndexNullsAllow
    nClsDb.CreateIndex "QType", "tblQuestion", "QType", False, adIndexNullsAllow

    nClsDb.CreateIndex "SubjectID", "tblScores", "SubjectID", False, adIndexNullsAllow
    nClsDb.CreateIndex "OfficialNo", "tblScores", "OfficialNo", False, adIndexNullsAllow
    nClsDb.CreateIndex "ScoreDate", "tblScores", "ScoreDate", False, adIndexNullsAllow

    nClsDb.CreateIndex "Rank", "tblUsers", "Rank", False, adIndexNullsAllow

    Set nClsDb = Nothing

End Sub

