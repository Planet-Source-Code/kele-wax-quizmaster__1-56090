VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.

Option Explicit

Private cat As ADOX.Catalog, tbl As ADOX.Table
Private col As ADOX.Column, idx As ADOX.Index
Private key As ADOX.key
Private fso As New Scripting.FileSystemObject
Private m_CN As String
Public Function CreateColumn(ByVal TabNam As String, ByVal ColNam As String, ByVal Typ As DataTypeEnum, _
                             Optional ByVal Size As Integer = 255, Optional ByVal AutoInc As Integer = 0, Optional ByVal Nullable As Integer = 0) As Boolean

    On Error GoTo CreateColumn_Error
    GetCatalog
    Set tbl = cat.Tables(TabNam)
    Set col = New ADOX.Column

    With col
        .Name = ColNam
        .Type = Typ
        If .Type >= adVarChar Then .DefinedSize = Size
        Set .ParentCatalog = cat
        If .Type = adInteger Then .Properties("Autoincrement") = (AutoInc = 1)
        .Properties("Nullable") = (Nullable = 1)
    End With
    tbl.Columns.Append col
    CreateColumn = True

    Exit Function

CreateColumn_Error:
    MsgBox Err.Description

End Function

Public Function CreateConstraintKey(ByVal KeyNam As String, _
                                    ByVal TabNam As String, ByVal ColNam As String, _
                                    ByVal VTabNam As String, ByVal VColNam As String)

    GetCatalog
    On Error GoTo CreateConstraintKey_Error
    Set key = New ADOX.key

    With key
        .Name = KeyNam
        .Type = adKeyForeign
        .RelatedTable = VTabNam
        .UpdateRule = adRICascade
        .DeleteRule = adRINone
        key.Columns.Append ColNam
        key.Columns(ColNam).RelatedColumn = VColNam
    End With
    Set tbl = cat.Tables(TabNam)
    tbl.Keys.Append key
CreateConstraintKey_Error:

End Function

Public Function CreateDB(DBName As String, Optional KillOldDb As Boolean = True) As Boolean

    On Error GoTo CreateDB_Error
    CreateDB = False
    If fso.FileExists(DBName) Then
        If KillOldDb Then
            fso.DeleteFile DBName
          Else
            Exit Function
        End If
    End If
    Set cat = New ADOX.Catalog
    cat.Create m_CN
    CreateDB = True

    Exit Function

CreateDB_Error:
    MsgBox Err.Description

End Function

Public Function CreateIndex(ByVal IDXNam As String, ByVal TabNam As String, _
                            ByVal ColNam As String, _
                            nodoubles As Boolean, Indexnull As AllowNullsEnum)

    GetCatalog
    On Error GoTo CreateIndex_Error
    Set idx = New ADOX.Index

    With idx
        .Name = IDXNam
        .IndexNulls = Indexnull
        .PrimaryKey = False
        .Unique = nodoubles
        .Columns.Append ColNam
    End With
    Set tbl = cat.Tables(TabNam)
    tbl.Indexes.Append idx
CreateIndex_Error:

End Function

Public Function CreatePrimaryKey(ByVal TabNam As String, ByVal ColNam As String)

    On Error GoTo CreatePrimaryKey_Error
    GetCatalog
    Set idx = New ADOX.Index

    With idx
        .Name = "PrimaryKey"
        .IndexNulls = adIndexNullsDisallow
        .PrimaryKey = True
        .Unique = True
        .Columns.Append ColNam
    End With
    Set tbl = cat.Tables(TabNam)
    tbl.Indexes.Append idx
CreatePrimaryKey_Error:

End Function

Public Function CreateTable(ByVal TabNam As String) As Boolean

    On Error GoTo CreateTable_Error
    GetCatalog
    Set tbl = New ADOX.Table
    tbl.Name = TabNam
    cat.Tables.Append tbl
    CreateTable = True

    Exit Function

CreateTable_Error:
    MsgBox Err.Description

End Function

Public Property Let DBConnection(Valu As String)

    m_CN = Valu

End Property

Public Property Get DBConnection() As String

    DBConnection = m_CN

End Property

Private Sub GetCatalog()

    If cat Is Nothing Then
        Set cat = New ADOX.Catalog
        cat.ActiveConnection = m_CN
    End If

End Sub

