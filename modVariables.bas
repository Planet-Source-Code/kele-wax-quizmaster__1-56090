Attribute VB_Name = "modVariables"
Public Type tTotalQuests
    tqOptions As Integer
    tqChoice As Integer
    tqWrite As Integer
End Type


'********* ADO VARIABLE ***********'

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public AddScore As New ADODB.Recordset
Public Find As New ADODB.Recordset



'********* INTEGER VARIABLE ***********'

Public Quest_No As Integer
Public i As Integer         'To keep tract total questions asked randomly
Public num As Integer       'To keep tract of the number of questions to be asked
Public Incre As Integer     ' Counter variable for random question generation
Public j As Integer         'To keep tract of question number for awarding marks
Public total As Integer     ' To keep tract of Marks scored
Public k As Integer         ' To keep tract of question number (User Anwser)
Public B As Integer         ' To keep tract of question number (Show Question Status)
Public D As Integer         'To number questions being displayed
Public Correct_Answer As Integer 'To keep tract of questions answered correctly
Public TotalQuests As tTotalQuests  'Number of questions to be asked

'********* STRING VARIABLE ***********'

Public Path As String
Public Comment As String * 100
Public Remark As String * 100

'********* BOOLEAN VARIABLE ***********'

Public AddQuestion As Boolean


