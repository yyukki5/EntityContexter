VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntityContexter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Entity Contexer -----------------------------------------------------------------
' Version   : 0.4.0
' Project   : https://github.com/yyukki5/EntityContexter
' Depend on : Microsoft ActiveX Data Objects X.X Library, Microsoft ADO Ext. 2.8 for DDL and Security, Microsoft Scripting RunTime
' Before Using : Need to check "Trust access to the VBA project object model"
' Function     : Scaffolding DB, Update, Insert, Delete
' ---------------------------------------------------------------------------------

Private Const VBA_CLASS_MODULE_ID As Long = 2
Private Const LOOP_MAX As Long = 2


'Private Const ENTIINFO_ENTITY As String = "Entity"
Private Const ENTIINFO_ENTITIES_NAME As String = "EntitiesName"
Private Const ENTIINFO_ENTITY_NAME As String = "EntityName"
Private Const ENTIINFO_PROPERTIES As String = "Properties"
Private Const ENTIINFO_PROPERTY_NAME As String = "PropertyName"
Private Const ENTIINFO_PROPERTY_TYPE As String = "PropertyType"
Private Const ENTIINFO_PROPERTY_FIELD_NAME As String = "PropertyFieldName"
Private Const ENTIINFO_DB_TABLE_NAME As String = "DbTableName"
Private Const ENTIINFO_DB_SQL_TABLE_NAME As String = "SQLTableName"
Private Const ENTIINFO_DB_FIELD_NAME As String = "DbFieldName"
Private Const ENTIINFO_DB_FIELD_TYPE As String = "DbFieldType"
Private Const ENTIINFO_DB_IS_AUTOINCREMENT As String = "DbIsAutoNumber"
Private Const ENTIINFO_DB_IS_PRIMARY_KEY As String = "DbIsPrimaryKey"
Private Const ENTIINFO_DB_IS_UNIQUE_KEY As String = "DbIsUniqueKey"
Private Const ENTIINFO_DB_IS_INDEX As String = "DbIsIndex"

Private Const DBINFO_ENTITY_NAME As String = "EcEntityName"
Private Const DBINFO_ENTITIES_NAME As String = "EcEntitiesName"
Private Const DBINFO_TABLE_NAME As String = "DbTableName"
Private Const DBINFO_SQL_TABLE_NAME As String = "SQLTableName"
Private Const DBINFO_PROPERTIES As String = "Properties"
Private Const DBINFO_PROPERTY_NAME As String = "PropertyName"
Private Const DBINFO_FIELD_NAME As String = "DbFieldName"
Private Const DBINFO_FIELD_TYPE As String = "DbFieldType"
Private Const DBINFO_IS_AUTONUMBER As String = "DbIsAutoNumber"
Private Const DBINFO_IS_PRIMARY_KEY As String = "DbIsPrimaryKey"
Private Const DBINFO_IS_UNIQUE_KEY As String = "DbIsUniqueKey"
Private Const DBINFO_INDEX As String = "DbIsPrimaryKey"



Private Const KEY_ENTITY_NAME As String = "EntityName"
Private Const KEY_ENTITIES_NAME As String = "EntitiesName"
Private Const KEY_ENTITY_ADOX_TABLE As String = "ADOX.Table"
Private Const KEY_DB_TABLE_NAME As String = "DBTableName"
Private Const KEY_SQL_TABLE_NAME As String = "SqlTableName"
Private Const KEY_TABLE_PRIMARY_KEY As String = "PrimaryKey"

Private Const KEY_ENTITY_PROPERTIES As String = "Properties"
Private Const KEY_ENTITY_PROPERTY_NAME As String = "PropertyName"
Private Const KEY_ENTITY_FIELD_NAME As String = "EntityFieldName"
Private Const KEY_ENTITY_PROPERTY_TYPE As String = "EntityPropertyType"
Private Const KEY_ENTITY_IS_PROPERTY_RELATED As String = "RelatedProperty"
Private Const KEY_DB_ADOX_COLUMN As String = "ADOX.Column"
Private Const KEY_ENTITY_EXIST_FOREIGNKEY As String = "ExistForeignKey"
Private Const KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY As String = "ExistInverseForeignKey"
Private Const KEY_ENTITY_FOREIGNKEY As String = "ForeignKey"
Private Const KEY_ENTITY_FOREIGN_TABLE As String = "ForeignTable"
Private Const KEY_ENTITY_INVERSE_FOREIGNKEY As String = "InverseForeignKey"
Private Const KEY_ENTITY_INVERSE_FOREIGN_TABLE As String = "InverseForeignTable"
Private Const KEY_DB_FIELD_NAME As String = "DbFieldName"
Private Const KEY_DB_FIELD_TYPE As String = "DbFieldType"



Private Const KEY_DB_RELATED_TABLE_NAME As String = "RelationBaseTableName"
Private Const KEY_DB_RELATED_FIELD_NAME As String = "RelationBaseFieldName"
Private Const KEY_DB_BASE_KEY_TABLE_NAME As String = "RelationRelatedTableName"
Private Const KEY_DB_BASE_KEY_FIELD_NAME As String = "RelationRelatedFieldName"


Private Const KEY_DB_AUTOINCREMENT As String = "Autoincrement"


Private cn_ As ADODB.Connection
Private catalog_ As ADOX.Catalog
Private isExportFiles_ As Boolean
Private schemaName_ As String
Private connectionString_ As String

Private fetchQuery_ As String
Private fetchEntityName_ As String
Private fetchEntityNames_ As Collection
Private lastSubQueryName_ As String
Private fetchFieldNames_ As New Dictionary

Private Sub Class_Initialize()
    Set cn_ = New ADODB.Connection
    Set catalog_ = New ADOX.Catalog
    isExportFiles_ = True
End Sub

Public Function Initialize(connectionString As String) As EntityContexter
    
    connectionString_ = connectionString
    cn_.connectionString = connectionString
    
    Set Initialize = Me
End Function

Private Sub Class_Terminate()
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Sub


Public Function ScaffoldDataBese(Optional usingAttribute As Boolean = True, _
                                 Optional schema_name As String = "dbo", _
                                 Optional dbcontext_name As String = "DbContext", _
                                 Optional find_word As String = "" _
                                 ) As Collection
    Dim table_name As Collection, tableNames As New Collection, entityInfos As New Collection
    
    ' setting
    OpenDb
    isExportFiles_ = usingAttribute
    schemaName_ = schema_name
    
    ' get table names
    Set tableNames = GetTableNameFromDb(find_word)
    
    ' create class module code from ADOX.catalog by table name
    For Each table_name In tableNames
        Call entityInfos.Add(GetEntityInfoByTableName(table_name), table_name(ENTIINFO_DB_TABLE_NAME))
    Next
    
    ' Add relationship to class module code
'    Call SetEntityInfoRelatedColumns(entityInfos)
                
    ' create class module from table names
    For Each table_name In tableNames
        Call CreateEntityClassFromTableName(entityInfos(table_name(ENTIINFO_DB_TABLE_NAME)))
'        Debug.Print table_name(KEY_DB_TABLE_NAME) & ": done"
    Next
    ' to be implement:  create DbContext.cls
    Call CreateDbContext(entityInfos, dbcontext_name)
    Debug.Print "- Done -------"
    
    ' for test
    Set ScaffoldDataBese = entityInfos
End Function

Public Function CreateDbInfoBy( _
    Optional table_name As String, _
    Optional sql_table_name As String, _
    Optional entity_name As String, _
    Optional entities_name As String _
    )
    Dim info As Collection
    
    Set info = New Collection
    Call info.Add(table_name, DBINFO_TABLE_NAME)
    Call info.Add(sql_table_name, DBINFO_SQL_TABLE_NAME)
    Call info.Add(entity_name, DBINFO_ENTITY_NAME)
    Call info.Add(entities_name, DBINFO_ENTITIES_NAME)
    
    Set CreateDbInfoBy = info
End Function

Public Function CreateDbInfoPropertyBy( _
    Optional property_name As String, _
    Optional db_field_name As String, _
    Optional db_field_type As ADOX.DataTypeEnum = 0, _
    Optional auto_increment As Boolean = False, _
    Optional primary_key As Boolean = False, _
    Optional unique_key As Boolean = False _
    )
    Dim prop As Collection
    
    Set prop = New Collection
    Call prop.Add(property_name, DBINFO_PROPERTY_NAME)
    Call prop.Add(db_field_name, DBINFO_FIELD_NAME)
    Call prop.Add(db_field_type, DBINFO_FIELD_TYPE)
    Call prop.Add(auto_increment, DBINFO_IS_AUTONUMBER)
    Call prop.Add(primary_key, DBINFO_IS_PRIMARY_KEY)
    Call prop.Add(unique_key, DBINFO_IS_UNIQUE_KEY)
    
    Set CreateDbInfoPropertyBy = prop
End Function


Public Function SetQueryDbSet(db_info, entity_name) As String
    fetchEntityName_ = entity_name
    Set fetchEntityNames_ = New Collection
    Call fetchEntityNames_.Add(entity_name)
    fetchQuery_ = "select * from " & db_info(entity_name)(DBINFO_SQL_TABLE_NAME)
End Function

Public Function SetQuerySelect(db_info, ParamArray property_names())
    Dim p, field_name As String, joined_field_name As String
    For Each p In property_names
        field_name = db_info(fetchEntityName_)(DBINFO_PROPERTIES)(p)(DBINFO_FIELD_NAME)
        joined_field_name = p & ","
    Next
    joined_field_name = Left(joined_field_name, Len(joined_field_name) - 1)
    lastSubQueryName_ = lastSubQueryName_ & "T"
    fetchQuery_ = "select " & joined_field_name & " from (" & fetchQuery_ & ") " & lastSubQueryName_
End Function
Public Function SetQueryWhere(db_info, property_name As String, sql_operator As String, compare_value As String, Optional entity_name As String)
    entity_name = IIf(entity_name = "", fetchEntityName_, entity_name)
    lastSubQueryName_ = lastSubQueryName_ & "T"
    fetchQuery_ = "select * from (" & fetchQuery_ & ") " & lastSubQueryName_ & " Where " & lastSubQueryName_ & "." & db_info(entity_name)(DBINFO_PROPERTIES)(property_name)(DBINFO_FIELD_NAME) & " " & sql_operator & " " & compare_value
End Function

Public Function SetQueryOrderBy(db_info, property_name As String, Optional entity_name As String, Optional is_asc As Boolean) As String
    entity_name = IIf(entity_name = "", fetchEntityName_, entity_name)
    lastSubQueryName_ = lastSubQueryName_ & "T"
    fetchQuery_ = "select * from (" & fetchQuery_ & ") " & lastSubQueryName_ & " order by " & lastSubQueryName_ & "." & db_info(entity_name)(DBINFO_PROPERTIES)(property_name)(DBINFO_FIELD_NAME) & IIf(is_asc, " asc", " desc")
End Function
Public Function SetQueryThenBy(db_info, property_name As String, Optional entity_name As String, Optional is_asc As Boolean) As String
    entity_name = IIf(entity_name = "", fetchEntityName_, entity_name)
    fetchQuery_ = fetchQuery_ & "," & lastSubQueryName_ & "." & db_info(entity_name)(DBINFO_PROPERTIES)(property_name)(DBINFO_FIELD_NAME) & IIf(is_asc, " asc", " desc")
End Function

Public Function SetQueryInclude(db_info As Variant, entity_name As String, entity_property_name As String) As String
    Call fetchEntityNames_.Add(entity_property_name)


'    Dim table_name As String, related_entity_name As String, foregin_table_name As String, base_field_name As String, related_field_name As String
'    Dim key_infos As Dictionary, key_info As Collection, base_sql_table_name As String, related_sql_table_name As String, ctxt As String, coltxt As String
'
'    lastSubQueryName_ = lastSubQueryName_ & "T"
'
'    table_name = db_info(entity_name)(DBINFO_TABLE_NAME)
'    foregin_table_name = db_info(entity_property_name)(DBINFO_TABLE_NAME)
'
'
'    For Each key_info In GetRelations(table_name, foregin_table_name)
'        base_sql_table_name = db_info(entity_name)(DBINFO_SQL_TABLE_NAME)
'        base_field_name = key_info(KEY_DB_BASE_KEY_FIELD_NAME)
'        related_field_name = key_info(KEY_DB_RELATED_FIELD_NAME)
'
''        ctxt = db_info(entity_name)(DBINFO_SQL_TABLE_NAME) & "." & base_field_name & " = " & db_info(entity_property_name)(DBINFO_SQL_TABLE_NAME) & "." & related_field_name & " And "
'        ctxt = lastSubQueryName_ & "." & base_field_name & " = " & db_info(entity_property_name)(DBINFO_SQL_TABLE_NAME) & "." & related_field_name & " And "
'    Next
'    ctxt = Left(ctxt, Len(ctxt) - Len(" And "))
'    Call SetFieldNamesStringFromDbInfoBy(db_info, entity_name, entity_property_name)
'
'    fetchQuery_ = "select " & Join(fetchFieldNames_.Keys, ",") & " from (" & fetchQuery_ & ") " & lastSubQueryName_ & " left outer join " & foregin_table_name & " on " & ctxt
End Function


Private Function GetEntFromEnts(tbl As Collection, rels As Collection, db_info As Collection)
    Dim prop As Collection, prop_name As String, ents As New Collection, ent, v As Variant, rel
    
    Set ents = New Collection
    For Each ent In tbl
        Dim f As Boolean: f = True
        For Each rel In rels
            prop_name = rel(DBINFO_PROPERTY_NAME)
            v = rel("Value")
            If CallByName(ent, prop_name, VbGet) = v Then Else f = False
        Next
        If f Then Call ents.Add(ent)
    Next
    
    Set GetEntFromEnts = ents
End Function


Public Function FetchRecords(entity_name As String, DbInfo As Object, dbc As Object, cnt As Long, Optional qry As String)
    
    
    ' -----
    Dim n, entss As New Collection
    For Each n In fetchEntityNames_
        Call entss.Add(FetchRecordsSub((n), DbInfo, dbc, 0), n)
    Next
    
    Dim ents, ent
    Dim prop, property_name As String, p, ent_name, is_entities As Boolean
    Set ents = entss(entity_name)
    For Each ent In ents
        For Each prop In DbInfo(entity_name)(DBINFO_PROPERTIES)
            property_name = prop(DBINFO_PROPERTY_NAME)
            If prop(DBINFO_FIELD_NAME) = "" Then
                Set p = CallByName(ent, prop(DBINFO_PROPERTY_NAME), VbGet)
                ent_name = GetEntityNameFromDbInfoBy(DbInfo, TypeName(p))
                
                is_entities = False
                If ent_name <> "" Then is_entities = True Else ent_name = TypeName(p)
                If CollectionContain(fetchEntityNames_, ent_name) Then
                
                    Dim rels As Collection, rel As Collection, p1 As Collection, p2 As Collection
                    Set rels = GetRelations((DbInfo(entity_name)(DBINFO_TABLE_NAME)), (DbInfo(ent_name)(DBINFO_TABLE_NAME)))
                    
                    For Each rel In rels
                        Set p1 = GetPropByTableFeildName(DbInfo, (DbInfo(entity_name)(DBINFO_TABLE_NAME)), rel(KEY_DB_BASE_KEY_FIELD_NAME))
                        Set p2 = GetPropByTableFeildName(DbInfo, (DbInfo(ent_name)(DBINFO_TABLE_NAME)), rel(KEY_DB_RELATED_FIELD_NAME))
                        Call rel.Add(p2(DBINFO_PROPERTY_NAME), DBINFO_PROPERTY_NAME)
                        Call rel.Add(CallByName(ent, p1(DBINFO_PROPERTY_NAME), VbGet), "Value")
                    Next
                                    
                    Dim tbls As Collection
                    Set tbls = GetEntFromEnts(entss(ent_name), rels, DbInfo)
                    If is_entities Then
                        If (tbls.Count > 0) Then Call CallByName(CallByName(ent, property_name, VbGet), "Items", VbSet, tbls)
                    Else
                        If (tbls.Count > 0) Then Call CallByName(ent, property_name, VbSet, tbls.Item(1))
                    End If
                End If
            End If
        Next
    Next
    Set FetchRecords = ents
    
    
           
    ' -----
 
    
    
'
'    Dim rs As Recordset, Items As New Collection, T As Object, props As Collection, prop As Collection, p As Variant
'    Dim fkey, kyfld, kyfldv, foreign_table_name As String, ent_name As String, entities_name As String, property_name As String, qry1, tbls As Collection, rels As Collection, rel As Collection
'    Dim sqn, wtxt, IsEntities As Boolean
'
'
'    If cnt > LOOP_MAX Then GoTo errFetchQuery
'
'    Set props = DbInfo(entity_name)(DBINFO_PROPERTIES)
'
'    ' todo---
''    joined_field_name = GetFieldNamesStringFromDbInfoBy(DbInfo, entity_name)
'    If qry = "" Then qry = "select " & GetFieldNamesStringFromDbInfoBy(DbInfo, entity_name) & " from (" & fetchQuery_ & ") " & lastSubQueryName_ & "T"
'    'If qry = "" Then qry = fetchQuery_ ' todo
'    ' ---
'
'    On Error GoTo errFetchQuery
'    Set rs = FetchRecordSetByQuery(qry)
'    If rs.BOF And rs.EOF Then GoTo errFetchQuery
'    On Error GoTo 0
'
'    rs.MoveFirst
'    Do Until rs.EOF
'        Set T = dbc.CreateObjectBy(entity_name)
'        For Each prop In props
'            property_name = prop(DBINFO_PROPERTY_NAME)
'            If prop(DBINFO_FIELD_NAME) = "" Then
'                Set p = CallByName(T, prop(DBINFO_PROPERTY_NAME), VbGet)
'                ent_name = GetEntityNameFromDbInfoBy(DbInfo, TypeName(p))
'                IsEntities = False
'                If ent_name <> "" Then IsEntities = True Else ent_name = TypeName(p)
'
'                sqn = lastSubQueryName_ & "T": wtxt = ""
'                For Each rel In GetRelations((DbInfo(entity_name)(DBINFO_TABLE_NAME)), (DbInfo(ent_name)(DBINFO_TABLE_NAME)))
'                    kyfldv = rs.Fields(rel(KEY_DB_BASE_KEY_FIELD_NAME)).Value  ' todo convert type
'                    wtxt = wtxt & DbInfo(ent_name)(DBINFO_TABLE_NAME) & "." & rel(KEY_DB_RELATED_FIELD_NAME) & " = " & kyfldv & " And "
'                Next
'                wtxt = TrimLast(wtxt, " And ")
''                qry1 = "SELECT * FROM (" & fetchQuery_ & ") " & sqn & " WHERE " & wtxt
'                qry1 = "select * from " & DbInfo(ent_name)(DBINFO_TABLE_NAME) & " " & DbInfo(ent_name)(DBINFO_TABLE_NAME) & " WHERE " & wtxt
'
'                If IsEntities Then
'                    Set tbls = FetchRecords(ent_name, DbInfo, dbc, cnt + 1, (qry1))
'                    If (tbls Is Nothing) = False Then Call CallByName(CallByName(T, property_name, VbGet), "Items", VbSet, tbls)
'
'                Else
''                    Set rel = GetRelations((DbInfo(entity_name)(DBINFO_TABLE_NAME)), (DbInfo(TypeName(p))(DBINFO_TABLE_NAME)))
''                    kyfldv = rs.Fields(rel(1)(KEY_DB_BASE_KEY_FIELD_NAME)).Value
''                    qry1 = "SELECT * FROM (" & fetchQuery_ & ") " & lastSubQueryName_ & "T" & " WHERE " & lastSubQueryName_ & "T." & rel(1)(KEY_DB_BASE_KEY_FIELD_NAME) & " = " & kyfldv ' todo convert type
'
'                    Set tbls = FetchRecords(TypeName(p), DbInfo, dbc, cnt + 1, (qry1))
'                    If (tbls Is Nothing) = False Then Call CallByName(T, property_name, VbSet, tbls.Item(1))
'                End If
'            Else
'                Call CallByName(T, property_name, VbLet, rs.Fields(prop(DBINFO_FIELD_NAME)).Value)
'            End If
'        Next
'        Items.Add T
'        rs.MoveNext
'    Loop
'    rs.Close
'
'    Set FetchRecords = Items
'    Exit Function
'
'errFetchQuery:
'    Set FetchRecords = Nothing:     Exit Function
End Function

Private Function FetchRecordsSub(entity_name As String, DbInfo As Object, dbc As Object, cnt As Long, Optional qry As String)
    Dim rs As Recordset, Items As New Collection, T As Object, props As Collection, prop As Collection, p As Variant
    Dim fkey, kyfld, kyfldv, foreign_table_name As String, ent_name As String, entities_name As String, property_name As String, qry1, tbls As Collection, rels As Collection, rel As Collection
           
    
    On Error GoTo errFetchQuery
    qry = "select * from " & DbInfo(entity_name)(DBINFO_SQL_TABLE_NAME)
    Set rs = FetchRecordSetByQuery(qry)
    If rs.BOF And rs.EOF Then GoTo errFetchQuery
    On Error GoTo 0
    
    Set props = DbInfo(entity_name)(DBINFO_PROPERTIES)
    rs.MoveFirst
    Do Until rs.EOF
        Set T = dbc.CreateObjectBy(entity_name)
        For Each prop In props
            property_name = prop(DBINFO_PROPERTY_NAME)
            If prop(DBINFO_FIELD_NAME) = "" Then
            Else
                Call CallByName(T, property_name, VbLet, rs.Fields(prop(DBINFO_FIELD_NAME)).Value)
            End If
        Next
        Items.Add T
        rs.MoveNext
    Loop
    rs.Close
    
    Set FetchRecordsSub = Items
    Exit Function
    
errFetchQuery:
    Set FetchRecordsSub = Nothing:     Exit Function
End Function



Private Function DoesPrimaryKeyHaveValue(rec As Object, dbInfos As Object) As Boolean
    Dim tbl As ADOX.Table, ky As ADOX.Key, col As ADOX.Column, DbInfo As Collection
    DoesPrimaryKeyHaveValue = True
    Set DbInfo = dbInfos(TypeName(rec))
    Set tbl = catalog_.Tables(DbInfo(KEY_DB_TABLE_NAME))
    For Each ky In tbl.Keys
        If ky.Type = adKeyPrimary Then
            For Each col In ky.Columns
                If IsNullOrEmptyOrZeroOrWhiteSpace(CallByName(rec, GetPropertyNameByDbTableFieldName(dbInfos, tbl.Name, col.Name), VbGet)) Then
                    DoesPrimaryKeyHaveValue = False
                End If
            Next
        End If
    Next
End Function

Private Function IsNullOrEmptyOrZeroOrWhiteSpace(v As Variant) As Boolean
    IsNullOrEmptyOrZeroOrWhiteSpace = True
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If vbInteger <= VarType(v) And VarType(v) <= vbCurrency Then If (v) = 0 Then Exit Function
    If VarType(v) = vbString Then If (v) = "" Then Exit Function
    IsNullOrEmptyOrZeroOrWhiteSpace = False
End Function

Private Function IsEntities(ent, prop As Collection, db_info) As Boolean
    Dim p, ent_name
    Set p = CallByName(ent, prop(DBINFO_PROPERTY_NAME), VbGet)
    ent_name = GetEntityNameFromDbInfoBy(db_info, TypeName(p))
    If ent_name <> "" Then IsEntities = True
End Function
Private Function GetEntityName(ent, prop As Collection, db_info) As String
    Dim p, ent_name
    Set p = CallByName(ent, prop(DBINFO_PROPERTY_NAME), VbGet)
    ent_name = GetEntityNameFromDbInfoBy(db_info, TypeName(p))
    If ent_name <> "" Then Else ent_name = TypeName(p)
    GetEntityName = ent_name
End Function

Public Function Add(rec As Object, DbInfo As Object)
    Dim rs As New ADODB.Recordset, prop As Collection
                                            
On Error GoTo errOpen
    OpenDb
    cn_.CursorLocation = adUseClient
    
    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) = "" Then
            Dim is_entities As Boolean, ent_name
            ent_name = GetEntityName(rec, prop, DbInfo)
            is_entities = IsEntities(rec, prop, DbInfo)
            
            If is_entities Then
                Dim r As Object
                For Each r In CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet)
                    
                    Call Add(r, DbInfo)
                Next
            Else
    '            If DoesPrimaryKeyHaveValue(rec, dbInfo) Then If DoesUniqueKeyHaveValue(rec, dbInfo) Then Exit Function
                If DoesPrimaryKeyHaveValue(rec, DbInfo) Then Exit Function
                Call Add(CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet), DbInfo)
                
                Dim v
                v = CallByName(CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet), _
                    GetPropByTableFeildName(DbInfo, prop(KEY_ENTITY_INVERSE_FOREIGN_TABLE), prop(KEY_ENTITY_INVERSE_FOREIGNKEY))(KEY_ENTITY_PROPERTY_NAME), VbGet)
                    
                Call CallByName(rec, GetPropByTableFeildName(prop, prop(KEY_DB_BASE_KEY_TABLE_NAME), prop(KEY_DB_BASE_KEY_FIELD_NAME))(KEY_ENTITY_PROPERTY_NAME), VbLet, v)
            End If
        End If
    Next
                
    If HasPrimaryKey(rec, DbInfo) = False Then If ExistUniqueRecord(rec, DbInfo) Then Call Err.Raise(10003, "EntityContexter.Add()", "Record is already exist")
    
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call rs.Open(Source:=DbInfo(TypeName(rec))(KEY_SQL_TABLE_NAME), ActiveConnection:=cn_, LockType:=adLockOptimistic)
    Call rs.AddNew(GetFieldsListByEntity(rec, DbInfo), GetValuesListByEntity(rec, DbInfo))
    Call SetAutoincrementFieldsValues(rs, rec, DbInfo)
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
errOpen:
    If cn_.State <> adStateClosed Then Call cn_.Close
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Add()", Err.Description
End Function

Private Function GetValuesListByEntity(rec As Object, DbInfo As Object)
    Dim prop As Collection, col As New Collection
    
    For Each prop In DbInfo(TypeName(rec))(KEY_ENTITY_PROPERTIES)
        If prop(KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(KEY_DB_AUTOINCREMENT) = False Then col.Add CallByName(rec, prop(KEY_ENTITY_PROPERTY_NAME), VbGet)
        End If
    Next
    GetValuesListByEntity = CollectionToArray(col)

End Function

Private Function GetFieldsListByEntity(rec As Object, DbInfo As Object)
    Dim prop As Collection, col As New Collection


    For Each prop In DbInfo(TypeName(rec))(KEY_ENTITY_PROPERTIES)
        If prop(KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(KEY_DB_AUTOINCREMENT) = False Then col.Add prop(KEY_DB_FIELD_NAME)
        End If
    Next
    GetFieldsListByEntity = CollectionToArray(col)

End Function

'todo change Update()
Public Function Update(rec As Object, DbInfo As Object)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    cn_.CursorLocation = adUseClient
    If HasPrimaryKey(rec, DbInfo) = False Then Call Err.Raise(10001, "EntityContexter.Update()", "Table has no primary key")
    If ExistUniqueRecord(rec, DbInfo) = False Then Call Err.Raise(10002, "EntityContexter.Update()", "Not found records")
    
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call cn_.Execute(CommandText:=CreateUpdateQuery(rec, DbInfo))
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
errOpen:
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Update()", Err.Description
End Function

'todo change Remove()
Public Function Remove(rec As Object, DbInfo As Object)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    cn_.CursorLocation = adUseClient
    If HasPrimaryKey(rec, DbInfo) = False Then Call Err.Raise(10001, "EntityContexter.Remove()", "Table has no primary key")
    If ExistUniqueRecord(rec, DbInfo) = False Then Call Err.Raise(10002, "EntityContexter.Remove()", "Not found records")
    
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call cn_.Execute(CommandText:=CreateDeleteQuery(rec, DbInfo))
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
errOpen:
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Remove()", Err.Description
End Function

Private Sub MoveToRecord(rs As ADODB.Recordset, rec As Object, DbInfo As Collection)
    Dim index_field As String, index_value, prop As Collection

    index_field = DbInfo(TypeName(rec))(KEY_TABLE_PRIMARY_KEY)
    For Each prop In DbInfo(TypeName(rec))(KEY_ENTITY_PROPERTIES)
        If prop(KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(KEY_DB_FIELD_NAME) = index_field Then index_value = CallByName(rec, prop(KEY_ENTITY_PROPERTY_NAME), VbGet)
        End If
    Next
    
    rs.MoveFirst
    Do Until rs.EOF
        If rs.Fields(index_field).Value = index_value Then Exit Do
        rs.MoveNext
    Loop
End Sub

Private Function SetAutoincrementFieldsValues(rs As ADODB.Recordset, rec As Object, DbInfo As Object)
    Dim prop As Collection

    For Each prop In DbInfo(TypeName(rec))(KEY_ENTITY_PROPERTIES)
        If prop(KEY_DB_AUTOINCREMENT) = True Then _
            Call CallByName(rec, prop(KEY_ENTITY_PROPERTY_NAME), VbLet, rs.Fields(prop(KEY_DB_FIELD_NAME)).Value)
    Next
End Function

Private Function HasPrimaryKey(rec As Object, DbInfo As Object) As Boolean
    Dim ind As ADOX.Index
    
    Set catalog_.ActiveConnection = cn_
    For Each ind In catalog_.Tables(DbInfo(TypeName(rec))(DBINFO_TABLE_NAME)).Indexes
        If ind.PrimaryKey Then HasPrimaryKey = True: Exit Function
    Next
       
End Function

Private Function ExistUniqueRecord(rec As Object, dbInfos As Object) As Boolean
    Dim qry As String
    qry = "select * from " & dbInfos(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & GetQueryPrimaryWhere(rec, dbInfos)
    ExistUniqueRecord = cn_.Execute(CommandText:=qry).RecordCount = 1
End Function

Private Function GetQueryPrimaryWhere(rec As Object, dbInfos As Object) As String
    Dim qry As String, ind As ADOX.Index, col As ADOX.Column
    
    Set catalog_.ActiveConnection = cn_
    qry = " Where "
    For Each ind In catalog_.Tables(dbInfos(TypeName(rec))(DBINFO_TABLE_NAME)).Indexes
        If ind.PrimaryKey Then
            For Each col In ind.Columns
                qry = qry & "[" & col.Name & "] = " & ConvertValueToSqlQuery(GetDbInfoPropertyByFieldName(col.Name, TypeName(rec), dbInfos)(DBINFO_PROPERTY_NAME), rec, dbInfos) & ", "
            Next
        End If
    Next
    qry = Left(qry, Len(qry) - 2)
    
    GetQueryPrimaryWhere = qry
End Function

Private Function GetPropByTableFeildName(db_infos As Collection, table_name As String, field_name As String) As Collection
    Dim db_info, props, prop
    For Each db_info In db_infos
        If db_info(DBINFO_TABLE_NAME) = table_name Then
            For Each prop In db_info(DBINFO_PROPERTIES)
                If prop(DBINFO_FIELD_NAME) = field_name Then
                    Set GetPropByTableFeildName = prop
                End If
            Next
        End If
    Next
End Function



Private Function CreateDeleteQuery(rec As Object, DbInfo As Object) As String
    Dim table_name As String, key_prop_name As String, key_field_name As String, key_value As Variant
    
    ' todo update to delete related table's record by using inner join.
    CreateDeleteQuery = "Delete From " & DbInfo(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & GetQueryPrimaryWhere(rec, DbInfo)
    
End Function

Private Function CreateUpdateQuery(rec As Object, DbInfo As Object) As String
    Dim table_name As String, key_prop_name As String, key_field_name As String, key_value As Variant, qry As String, prop As Collection
    
    qry = "Update " & DbInfo(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & " set "
    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) = "" Then
        Else
            If prop(DBINFO_IS_AUTONUMBER) = False Then _
                qry = qry & "[" & prop(DBINFO_FIELD_NAME) & "] = " & ConvertValueToSqlQuery((prop(DBINFO_PROPERTY_NAME)), rec, DbInfo) & ", "
        End If
    Next
    qry = Left(qry, Len(qry) - 2) & GetQueryPrimaryWhere(rec, DbInfo)
    
    CreateUpdateQuery = qry
End Function

Private Function CreateInsertQuery(rec As Object, DbInfo As Object) As String
    Dim qry As String, field_names As String, field_values As String, prop As Collection
        
    qry = "Insert into " & DbInfo(TypeName(rec))(DBINFO_SQL_TABLE_NAME)
    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) = "" Then
        Else
            If prop(DBINFO_IS_AUTONUMBER) = False Then
                field_names = field_names & prop(DBINFO_FIELD_NAME) & ", "
                field_values = field_values & ConvertValueToSqlQuery((prop(DBINFO_PROPERTY_NAME)), rec, DbInfo) & ", "
            End If
        End If
    Next
    field_names = Left(field_names, Len(field_names) - 2)
    field_values = Left(field_values, Len(field_values) - 2)
    
    qry = qry & " (" & field_names & ") Values (" & field_values & ")"
    CreateInsertQuery = qry
End Function

Private Function ConvertValueToSqlQuery(prop_name As String, rec As Object, DbInfo As Object) As String
    Dim v As Variant, field_type As ADOX.DataTypeEnum
    v = CallByName(rec, prop_name, VbGet)
    field_type = DbInfo(TypeName(rec))(DBINFO_PROPERTIES)(prop_name)(DBINFO_FIELD_TYPE)
    Select Case VarType(v)
        Case vbEmpty:  v = "Null"
        Case vbNull:   v = "Null"
        Case vbString: v = "'" & v & "'"
        Case vbDate
            Select Case field_type
                Case ADOX.DataTypeEnum.adDBDate: v = Format(v, "'yyyy-MM-dd'")
                Case ADOX.DataTypeEnum.adDBTime: v = Format(v, "'hh:mm:ss'")
                Case Else: v = Format(v, "'yyyy-MM-dd hh:mm:ss'")
            End Select
    End Select
    
    ConvertValueToSqlQuery = v
End Function


Private Function GetTableNameByEntity(rec As Object, DbInfo As Object) As String
    Dim table_name As String, info As Collection
    
    For Each info In DbInfo
        If info(DBINFO_ENTITY_NAME) = TypeName(rec) Then table_name = info(DBINFO_SQL_TABLE_NAME)
    Next
    
    GetTableNameByEntity = table_name
End Function

Private Function GetPrimaryKeyPropertyByEntity(rec As Object, DbInfo As Object) As String
    Dim prop As Collection
    
    For Each prop In DbInfo(TypeName(rec))(KEY_ENTITY_PROPERTIES)
        If prop(KEY_DB_FIELD_NAME) = DbInfo(TypeName(rec))(KEY_TABLE_PRIMARY_KEY) Then GetPrimaryKeyPropertyByEntity = prop(KEY_ENTITY_PROPERTY_NAME)
    Next
    
End Function

Private Function GetDbInfoPropertyByFieldName(field_name As String, entity_name As String, dbInfos As Object) As Collection
    Dim prop As Collection
    
    For Each prop In dbInfos(entity_name)(KEY_ENTITY_PROPERTIES)
        If prop(KEY_DB_FIELD_NAME) = field_name Then Set GetDbInfoPropertyByFieldName = prop
    Next
    
End Function

Private Property Get ProviderHasSchema() As Boolean
    ProviderHasSchema = InStr(cn_.PROVIDER, "SQLOLEDB") > 0
End Property



' Private Methods ----------------------------------------------------------------
Private Function OpenDb()
    If cn_.State <> adStateOpen Then cn_.Open: Set catalog_.ActiveConnection = cn_
End Function

Private Function GetTableNameFromDb(Optional find_word As String = "") As Collection
    Dim tableNames As New Collection, sch As ADODB.Recordset, table_names As Collection
    
On Error GoTo errLabel
    Call OpenDb
    Set sch = cn_.OpenSchema(adSchemaTables)
    
    sch.MoveFirst
    Do Until sch.EOF
        Set table_names = New Collection
        
        If find_word <> "" And InStr(sch.Fields("TABLE_NAME").Value, find_word) = 0 Then GoTo next_loop
        
        If ProviderHasSchema Then
            If sch.Fields("TABLE_SCHEMA").Value = schemaName_ Then
                Call table_names.Add(sch.Fields("TABLE_NAME").Value, ENTIINFO_DB_TABLE_NAME)
                Call table_names.Add(ConvertToSqlTableNameBy(sch.Fields("TABLE_NAME").Value, schemaName_), ENTIINFO_DB_SQL_TABLE_NAME)
                            
                Call tableNames.Add(table_names)
            End If
        Else
            If sch.Fields("TABLE_TYPE").Value = "TABLE" Or sch.Fields("TABLE_TYPE").Value = "VIEW" Then
                Call table_names.Add(sch.Fields("TABLE_NAME").Value, ENTIINFO_DB_TABLE_NAME)
                Call table_names.Add(ConvertToSqlTableNameBy(sch.Fields("TABLE_NAME").Value), ENTIINFO_DB_SQL_TABLE_NAME)
                
                Call tableNames.Add(table_names)
            End If
        End If
next_loop:
        sch.MoveNext
    Loop
    Set GetTableNameFromDb = tableNames
    Exit Function

errLabel:
    Err.Raise 10000, Err.Source & vbCrLf & "at EntityContexter.GetTableNameFromDb()", Err.Description
End Function

Private Function ConvertToSqlTableNameBy(table_name As String, Optional schema_name As String) As String
    ConvertToSqlTableNameBy = IIf(schema_name = "", "[" & table_name & "]", "[" & schema_name & "]." & "[" & table_name & "]")
End Function
'Private Function ConvertToTableNameBy(sql_table_name As String) As String
'    If InStr(sql_table_name, "].[") > 0 Then
'        ConvertToTableNameBy = Mid(sql_table_name, InStr(sql_table_name, "].[") + 2, Len(sql_table_name) - 1)
'    Else
'        ConvertToTableNameBy = Mid(sql_table_name, 2, Len(sql_table_name) - 2)
'    End If
'End Function

Private Function GetEntityInfoByTableName(table_name_info As Collection) As Collection
    Dim props As New Collection, prop As Collection, tbl As ADOX.Table, col As ADOX.Column, entity_info As New Collection
    Dim ky As ADOX.Key, table_name As String
    
    table_name = table_name_info(ENTIINFO_DB_TABLE_NAME)
    Set tbl = catalog_.Tables(table_name)
    For Each col In tbl.Columns
        Set prop = GetPropertyInfoByColumn(col)
        Call prop.Add(IsIndexKeyBy(tbl.Indexes, col), ENTIINFO_DB_IS_INDEX)
        Call prop.Add(IsPrimaryKeyBy(tbl.Indexes, col), ENTIINFO_DB_IS_PRIMARY_KEY)
        Call prop.Add(IsUniqueKeyBy(tbl.Indexes, col), ENTIINFO_DB_IS_UNIQUE_KEY)
        Call prop.Add(GetIsAutoNumberField(col, tbl.Name), ENTIINFO_DB_IS_AUTOINCREMENT)
        
        Call props.Add(prop, prop(ENTIINFO_PROPERTY_NAME))
    Next
    
    For Each col In tbl.Columns
        Call SetPropertyInfoOfRelatedTabel(props, tbl, col)
    Next

    Call entity_info.Add(table_name_info(ENTIINFO_DB_TABLE_NAME), ENTIINFO_DB_TABLE_NAME)
    Call entity_info.Add(table_name_info(ENTIINFO_DB_SQL_TABLE_NAME), ENTIINFO_DB_SQL_TABLE_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntityClassName(table_name), ENTIINFO_ENTITY_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntitesClassName(table_name), ENTIINFO_ENTITIES_NAME)
'    Call entity_info.Add(tbl, KEY_ENTITY_ADOX_TABLE)
    Call entity_info.Add(props, ENTIINFO_PROPERTIES)
        
    Set GetEntityInfoByTableName = entity_info
End Function

Private Function GetPropertyInfoByColumn(col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbFieldNameToClassPropertyName(col.Name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(col.Name)
    propertyTypeName = GetEntityPropertyTypeBy(col)

    Dim prop As New Collection
    Call prop.Add(propertyName, ENTIINFO_PROPERTY_NAME)
    Call prop.Add(propertyTypeName, ENTIINFO_PROPERTY_TYPE)
    Call prop.Add(propertyFieldName, ENTIINFO_PROPERTY_FIELD_NAME)
'    Call prop.Add(False, KEY_ENTITY_IS_PROPERTY_RELATED)
'    Call prop.Add(col, KEY_DB_ADOX_COLUMN)
    Call prop.Add(col.Name, ENTIINFO_DB_FIELD_NAME)
'    Call prop.Add(col.Type, ENTIINFO_DB_FIELD_TYPE)
    
    Set GetPropertyInfoByColumn = prop
End Function

Private Function SetPropertyInfoOfRelatedTabel(props As Collection, base_table As ADOX.Table, base_column As ADOX.Column) As Collection
    Dim tbl As ADOX.Table, ky As ADOX.Key, col As ADOX.Column, prop As Collection
    Dim rel
    
    ' todo
'    For Each tbl In catalog_.Tables
'        For Each rel In GetRelations(base_table.Name, tbl.Name)
'            If rel(KEY_DB_BASE_KEY_TABLE_NAME) = base_table.Name Then
'                Set prop = GetPropertyInfoOfRelatedEntities((rel(KEY_DB_RELATED_TABLE_NAME)))
'                Call props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
'            End If
'        Next
'    Next
    
    ' todo modify to upper
    If CanGetAdoxKey Then
        For Each tbl In catalog_.Tables
            If tbl.Name = base_table.Name Then
                For Each ky In tbl.Keys
                    If ky.Type = adKeyForeign Then
                        For Each col In ky.Columns
                            If col.Name = base_column.Name Then
                                If IsPrimaryKeyBy(catalog_.Tables(ky.RelatedTable).Indexes, catalog_.Tables(ky.RelatedTable).Columns(col.RelatedColumn)) Then
                                    Set prop = GetPropertyInfoOfRelatedEntity(ky.RelatedTable)
                                    Call props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
                                Else
                                    Set prop = GetPropertyInfoOfRelatedEntities(ky.RelatedTable)
                                    Call props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
                                End If
                            End If
                        Next
                    End If
                Next
            Else
                For Each ky In tbl.Keys
                    If ky.RelatedTable = base_table.Name And ky.Type = adKeyForeign Then
                        For Each col In ky.Columns
                            If col.RelatedColumn = base_column.Name Then
                                If IsPrimaryKeyBy(base_table.Indexes, base_column) Then
                                    Set prop = GetPropertyInfoOfRelatedEntities(tbl.Name)
                                    Call props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
                                Else
                                    Set prop = GetPropertyInfoOfRelatedEntity(tbl.Name)
                                    Call props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
                                End If
                            End If
                        Next

                    End If

                Next
            End If
        Next

    Else
        ' todo case of T-SQL
        Dim qry As String
        qry = "SELECT fk.CONSTRAINT_NAME AS ForeignKey, fk.TABLE_NAME AS ChildTable, fk.COLUMN_NAME AS ChildColumn, pk.TABLE_NAME AS ParentTable, pk.COLUMN_NAME AS ParentColumn " & _
            "FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS AS rc JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS fk " & _
            "OifN rc.CONSTRAINT_NAME = fk.CONSTRAINT_NAME JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS pk ON rc.UNIQUE_CONSTRAINT_NAME = pk.CONSTRAINT_NAME Order BY ChildTable, ForeignKey;"

    End If
End Function


Private Function GetRelations(base_table_name As String, related_table_name As String) As Collection
    Dim tbl As ADOX.Table, ky As ADOX.Key, col As ADOX.Column, key_infos As New Collection, key_info As New Collection
    Call OpenDb
    
    If CanGetAdoxKey Then
        For Each ky In catalog_.Tables(base_table_name).Keys
            If ky.Type = adKeyForeign And ky.RelatedTable = related_table_name Then
                For Each col In ky.Columns
                    Call key_info.Add(base_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
                    Call key_info.Add(related_table_name, KEY_DB_RELATED_TABLE_NAME)
                    Call key_info.Add(col.Name, KEY_DB_BASE_KEY_FIELD_NAME)
                    Call key_info.Add(col.RelatedColumn, KEY_DB_RELATED_FIELD_NAME)
                    Call key_infos.Add(key_info, ky.Name)
                Next
            End If
        Next
        For Each tbl In catalog_.Tables
            If tbl.Name = related_table_name Then
                For Each ky In tbl.Keys
                    If ky.RelatedTable = base_table_name And ky.Type = adKeyForeign Then
                        For Each col In ky.Columns
                            Call key_info.Add(base_table_name, KEY_DB_RELATED_TABLE_NAME)
                            Call key_info.Add(related_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
                            Call key_info.Add(col.RelatedColumn, KEY_DB_BASE_KEY_FIELD_NAME)
                            Call key_info.Add(col.Name, KEY_DB_RELATED_FIELD_NAME)
                            Call key_infos.Add(key_info, ky.Name)
                        Next
                     End If
                Next
            End If
        Next
    Else
        Dim qry As String
        qry = "SELECT fk.CONSTRAINT_NAME AS ForeignKey, fk.TABLE_NAME AS ChildTable, fk.COLUMN_NAME AS ChildColumn, pk.TABLE_NAME AS ParentTable, pk.COLUMN_NAME AS ParentColumn " & _
            "FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS AS rc JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS fk " & _
            "OifN rc.CONSTRAINT_NAME = fk.CONSTRAINT_NAME JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS pk ON rc.UNIQUE_CONSTRAINT_NAME = pk.CONSTRAINT_NAME Order BY ChildTable, ForeignKey;"
        
        Dim rs, i As Long
        rs = cn_.Execute(qry).GetRows
        For i = LBound(rs) To UBound(rs)
            If rs(i, 1) = base_table_name And rs(i, 3) = related_table_name Then
                Call key_info.Add(rs(i, 1), KEY_DB_BASE_KEY_TABLE_NAME)
                Call key_info.Add(rs(i, 3), KEY_DB_RELATED_TABLE_NAME)
                Call key_info.Add(rs(i, 2), KEY_DB_BASE_KEY_FIELD_NAME)
                Call key_info.Add(rs(i, 4), KEY_DB_RELATED_FIELD_NAME)
                Call key_infos.Add(key_info, ky.Name)
            ElseIf rs(i, 1) = related_table_name And rs(i, 3) = base_table_name Then
                Call key_info.Add(rs(i, 3), KEY_DB_BASE_KEY_TABLE_NAME)
                Call key_info.Add(rs(i, 1), KEY_DB_RELATED_TABLE_NAME)
                Call key_info.Add(rs(i, 4), KEY_DB_BASE_KEY_FIELD_NAME)
                Call key_info.Add(rs(i, 2), KEY_DB_RELATED_FIELD_NAME)
                Call key_infos.Add(key_info, ky.Name)
            End If
        Next
        
    End If
    
    Set GetRelations = key_infos
End Function



Private Function GetPropertyInfoOfRelatedEntity(table_name As String)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbTableNameToEntityClassName(table_name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, ENTIINFO_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, ENTIINFO_PROPERTY_FIELD_NAME)
    Call prop.Add(propertyTypeName, ENTIINFO_PROPERTY_TYPE)
'    Call prop.Add(col, KEY_DB_ADOX_COLUMN)
    Call prop.Add("", ENTIINFO_DB_FIELD_NAME)
    
'    Call prop.Add(True, KEY_ENTITY_IS_PROPERTY_RELATED)
'    Call prop.Add(base_table_name, KEY_DB_RELATED_TABLE_NAME)
'    Call prop.Add(base_field_name, KEY_DB_RELATED_FIELD_NAME)
'    Call prop.Add(table_name, KEY_DB_BASE_KEY_TABLE_NAME)
'    Call prop.Add(related_field_name, KEY_DB_BASE_KEY_FIELD_NAME)
    
    Call prop.Add(False, ENTIINFO_DB_IS_INDEX)
    Call prop.Add(False, ENTIINFO_DB_IS_PRIMARY_KEY)
'   Call prop.Add(false, ENTIINFO_DB_IS_UNIQUE_KEY)
    Call prop.Add(False, ENTIINFO_DB_IS_AUTOINCREMENT)
    
    Set GetPropertyInfoOfRelatedEntity = prop
End Function

Private Function GetPropertyInfoOfRelatedEntities(related_table_name As String)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbTableNameToEntitesClassName(related_table_name)
    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(related_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, ENTIINFO_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, ENTIINFO_PROPERTY_FIELD_NAME)
    Call prop.Add(propertyTypeName, ENTIINFO_PROPERTY_TYPE)
'    Call prop.Add(col, KEY_DB_ADOX_COLUMN)
    Call prop.Add("", ENTIINFO_DB_FIELD_NAME)
    
'    Call prop.Add(True, KEY_ENTITY_IS_PROPERTY_RELATED)
'    Call prop.Add(base_table_name, KEY_DB_RELATED_TABLE_NAME)
'    Call prop.Add(base_field_name, KEY_DB_RELATED_FIELD_NAME)
'    Call prop.Add(related_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
'    Call prop.Add(related_field_name, KEY_DB_BASE_KEY_FIELD_NAME)
   
    Call prop.Add(False, ENTIINFO_DB_IS_INDEX)
    Call prop.Add(False, ENTIINFO_DB_IS_PRIMARY_KEY)
'   Call prop.Add(false, ENTIINFO_DB_IS_UNIQUE_KEY)
    Call prop.Add(False, ENTIINFO_DB_IS_AUTOINCREMENT)
    
    Set GetPropertyInfoOfRelatedEntities = prop
End Function


Private Function SetEntityInfoRelatedColumns(ByRef entityInfos As Collection)
    Dim EntityInfo As Collection, col As ADOX.Column, ky As ADOX.Key, tbl As ADOX.Table, prop As Collection
    
    For Each EntityInfo In entityInfos
        Set tbl = EntityInfo(KEY_ENTITY_ADOX_TABLE)
        For Each ky In tbl.Keys
            If ky.Type <> adKeyForeign Then
            Else
                For Each col In ky.Columns
                    Set prop = GetPropertyInfoByInverseProperty(ky.RelatedTable, col)
                    Call EntityInfo(KEY_ENTITY_PROPERTIES).Add(prop, prop(KEY_ENTITY_PROPERTY_NAME))
                    Set prop = GetPropertyInfoByForeignKey(tbl.Name, col)
                    Call entityInfos(ky.RelatedTable).Item(KEY_ENTITY_PROPERTIES).Add(prop, prop(KEY_ENTITY_PROPERTY_NAME))
                Next
            End If
        Next
    Next
    
End Function

Private Function GetPropertyInfoByForeignKey(key_table_name As String, col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertDbTableNameToEntitesClassName(key_table_name)
    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(key_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(True, KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col, KEY_DB_ADOX_COLUMN)
    Call prop.Add(col.Name, KEY_DB_FIELD_NAME)
    Call prop.Add(col.RelatedColumn, KEY_ENTITY_FOREIGNKEY)
    Call prop.Add(key_table_name, KEY_ENTITY_FOREIGN_TABLE)
    Call prop.Add("", KEY_ENTITY_INVERSE_FOREIGNKEY)
    Call prop.Add("", KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    
    Set GetPropertyInfoByForeignKey = prop
End Function

Private Function GetPropertyInfoByInverseProperty(key_table_name As String, col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbTableNameToEntityClassName(key_table_name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(key_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(True, KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col.RelatedColumn, KEY_DB_FIELD_NAME)
    Call prop.Add("", KEY_ENTITY_FOREIGNKEY)
    Call prop.Add("", KEY_ENTITY_FOREIGN_TABLE)
    Call prop.Add(col.Name, KEY_ENTITY_INVERSE_FOREIGNKEY)
    Call prop.Add(key_table_name, KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    Call prop.Add(col, KEY_DB_ADOX_COLUMN)
    
    Set GetPropertyInfoByInverseProperty = prop
End Function


Private Function GetIsAutoNumberField(col As ADOX.Column, table_name As String) As Boolean
    Dim qry As String
    
    If ProviderHasSchema Then
        qry = "SELECT c.name AS ColumnName " & "FROM sys.columns c " & "INNER JOIN sys.tables t ON c.object_id = t.object_id " & "INNER JOIN sys.schemas s ON t.schema_id = s.schema_id " & _
             "WHERE t.name = '" & table_name & "' " & "AND s.name = '" & schemaName_ & "' " & "AND c.is_identity = 1;"
        GetIsAutoNumberField = Replace(cn_.Execute(qry).GetString(RowDelimeter:=","), ",", "") = col.Name
    Else
        GetIsAutoNumberField = col.Properties(KEY_DB_AUTOINCREMENT)
    End If
    
End Function


'Private Function GetPropertyInfoByForeignKey(key_table_name As String, col As ADOX.Column)
'    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
'
'    propertyName = ConvertDbTableNameToEntitesClassName(key_table_name)
'    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(key_table_name)
'    propertyTypeName = propertyName
'
'    Dim prop As New Collection
'    Call prop.Add(propertyName, KEY_ENTITY_PROPERTY_NAME)
'    Call prop.Add(propertyFieldName, KEY_ENTITY_FIELD_NAME)
'    Call prop.Add(propertyTypeName, KEY_ENTITY_PROPERTY_TYPE)
'    Call prop.Add(True, KEY_ENTITY_IS_PROPERTY_RELATED)
'    Call prop.Add(col, KEY_DB_ADOX_COLUMN)
'    Call prop.Add(col.Name, KEY_DB_FIELD_NAME)
'    Call prop.Add(col.RelatedColumn, KEY_ENTITY_FOREIGNKEY)
'    Call prop.Add(key_table_name, KEY_ENTITY_FOREIGN_TABLE)
'    Call prop.Add("", KEY_ENTITY_INVERSE_FOREIGNKEY)
'    Call prop.Add("", KEY_ENTITY_INVERSE_FOREIGN_TABLE)
'
'    Set GetPropertyInfoByForeignKey = prop
'End Function


Private Function CreateDbContext(entity_Infos As Collection, dbcontext_name As String)
    Dim codeDbcClass As String
    
    codeDbcClass = GetDbContextCode(entity_Infos, dbcontext_name)
    If isExportFiles_ Then
        Call AddClassModuleFromFile("DBContext", codeDbcClass)
    Else
    End If

End Function

Private Function CreateEntityClassFromTableName(entityInfos As Collection)
    Dim EntityClass As Object, EntitiesClass As Object, codeEntityClass As String, codeEntitiesClass As String
        
    codeEntityClass = GetEntityCode(entityInfos(ENTIINFO_ENTITY_NAME), entityInfos(ENTIINFO_PROPERTIES))
    codeEntitiesClass = GetEntitesCode(entityInfos(ENTIINFO_ENTITIES_NAME), entityInfos(ENTIINFO_ENTITY_NAME))
    
    If isExportFiles_ Then
        ' option1 using write file and import file. to using attribute key word in class module.
        Call AddClassModuleFromFile(entityInfos(ENTIINFO_ENTITY_NAME), codeEntityClass)
        Call AddClassModuleFromFile(entityInfos(ENTIINFO_ENTITIES_NAME), codeEntitiesClass)
        
    Else
        ' option2 using CodeModule.Insertlines()
        ' only force update (delete all and re-insert)
        Set EntityClass = AddClassModule(entityInfos(ENTIINFO_ENTITY_NAME))
        Call EntityClass.CodeModule.DeleteLines(1, EntityClass.CodeModule.CountOfLines)
        Call EntityClass.CodeModule.InsertLines(1, codeEntityClass)
        
        Set EntitiesClass = AddClassModule(entityInfos(ENTIINFO_ENTITIES_NAME))
        Call EntitiesClass.CodeModule.DeleteLines(1, EntitiesClass.CodeModule.CountOfLines)
        Call EntitiesClass.CodeModule.InsertLines(1, codeEntitiesClass)
    End If
End Function

Private Function AddClassModuleFromFile(className As String, codes As String)
    Dim classFilePath As String
    classFilePath = ThisWorkbook.Path & "\" & className & ".cls"
    Call CreateClassModuleFile(classFilePath, codes)
    Call ImportClassModuleFile(classFilePath, className)
    Call DeleteClassModuleFile(classFilePath)
End Function

Private Function DeleteClassModuleFile(classFilePath)
    Kill classFilePath
End Function

Private Function CreateClassModuleFile(classFilePath, codes)
    Dim fso As Object, txtFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(classFilePath, True)
    txtFile.Write codes
End Function

Private Function ImportClassModuleFile(classFilePath, className)
    Dim vbc As Object
    With Application.VBE.ActiveVBProject
        For Each vbc In .VBComponents
            If vbc.Name = className Then .VBComponents.Remove .VBComponents(className)
        Next
        .VBComponents.Import classFilePath
    End With
End Function

Private Function AddClassModule(className As String) As Object
    Dim clsModule As Object

    With Application.VBE.ActiveVBProject
        Set clsModule = GetVBComponent(className)
        If clsModule Is Nothing Then
            Set clsModule = .VBComponents.Add(VBA_CLASS_MODULE_ID)
            clsModule.Name = className
        End If
    End With
    
    Set AddClassModule = clsModule
End Function

Private Function GetVBComponent(componentName As String) As Object
    Dim c
    
    Set GetVBComponent = Nothing
    With Application.VBE.ActiveVBProject
        For Each c In .VBComponents
            If c.Name = componentName Then
                Set GetVBComponent = c
                Exit Function
            End If
        Next
    End With
End Function

Private Function GetDbContextCode(infos As Collection, dbc_name As String) As String
    Dim texts As New Collection, str As String, txt, tmp As String
    Dim info As Collection, prop As Collection
    Const db_table_name As String = "DBTableName"
    
    On Error Resume Next
    With texts
        If isExportFiles_ Then
            .Add "VERSION 1.0 CLASS"
            .Add "BEGIN"
            .Add "  MultiUse = -1  'True"
            .Add "End"
            .Add "Attribute VB_Name = """ & dbc_name & """"
            .Add "Attribute VB_GlobalNameSpace = False"
            .Add "Attribute VB_Creatable = False"
            .Add "Attribute VB_PredeclaredId = False"
            .Add "Attribute VB_Exposed = False"
        End If
        .Add "Option Explicit"
        .Add ""
        .Add "Private ec_ As EntityContexter"
        .Add "Private entity_name_ As String"
        .Add ""
        .Add "Private Sub Class_Initialize()"
        .Add "    Set ec_ = New EntityContexter"
        .Add "End Sub"
        .Add ""
        .Add "Public Function Initialize(connection_string As String)"
        .Add ""
        .Add "    Call ec_.Initialize(connection_string)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add ""
        .Add "Public Function DbSet(entity_name As String) As " & dbc_name
        .Add "        Call ec_.SetQueryDbSet(DbInfo, entity_name):    entity_name_ = entity_name:     Set DbSet = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function Where(property_name As String, sql_operator As String, compare_value As String, Optional entity_name As String) As " & dbc_name
        .Add "        Call ec_.SetQueryWhere(DbInfo, property_name, sql_operator, compare_value, entity_name):      Set Where = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function Include(entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryInclude(DbInfo, entity_name_, entity_name):     Set Include = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function OrderBy(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryOrderBy(DbInfo, property_name, entity_name, true):   Set OrderBy = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function OrderByDescending(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryOrderBy(DbInfo, property_name, entity_name, false):   Set OrderByDescending = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenBy(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryThenBy(DbInfo, property_name, entity_name, false):   Set ThenBy = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenByDescending(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryThenBy(DbInfo, property_name, entity_name, false):   Set ThenByDescending = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ToList() As Variant"
        .Add "    Set ToList = ec_.FetchRecords(entity_name_, DbInfo, Me, 0)"
        .Add "End Function"
        .Add ""
        .Add ""
        .Add "Public Function Add(rec As Object)"
        .Add ""
        .Add "    Call ec_.Add(rec, DBInfo)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Update(rec As Object)"
        .Add ""
        .Add "    Call ec_.Update(rec, DBInfo)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Remove(rec As Object)"
        .Add ""
        .Add "    Call ec_.Remove(rec, DBInfo)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Private Property Get DbInfo() As Collection"
        .Add "    Dim infos As New Collection, info As Collection, props As Collection, prop As Collection"
        .Add ""
        For Each info In infos
            tmp = "    Set info = ec_.CreateDbInfoBy( "
            tmp = tmp & "  table_name    :=""" & info(ENTIINFO_DB_TABLE_NAME) & """"
            tmp = tmp & ", sql_table_name:=""" & info(ENTIINFO_DB_SQL_TABLE_NAME) & """"
            tmp = tmp & ", entity_name   :=""" & info(ENTIINFO_ENTITY_NAME) & """"
            tmp = tmp & ", entities_name :=""" & info(ENTIINFO_ENTITIES_NAME) & """"
'            tmp = tmp & ", primary_key   :=""" & info(ENTIINFO_DB_IS_PRIMARY_KEY) & """"
            tmp = tmp & ")"
            .Add tmp
            .Add ""
            .Add "    Set props = New Collection"
            For Each prop In info(ENTIINFO_PROPERTIES)
                tmp = "    Set prop = ec_.CreateDbInfoPropertyBy("
                tmp = tmp & "   property_name:=""" & prop(ENTIINFO_PROPERTY_NAME) & """"
                tmp = tmp & IIf(prop(ENTIINFO_DB_FIELD_NAME) <> "", ",  db_field_name:=""" & prop(ENTIINFO_DB_FIELD_NAME) & """", "")
'                tmp = tmp & ", entity_property_type:=""" & prop(KEY_ENTITY_PROPERTY_TYPE) & """"
                tmp = tmp & IIf(prop(ENTIINFO_DB_IS_AUTOINCREMENT), ", auto_increment:=" & prop(ENTIINFO_DB_IS_AUTOINCREMENT), "")
                tmp = tmp & IIf(prop(ENTIINFO_DB_IS_PRIMARY_KEY), ", primary_key:=" & prop(ENTIINFO_DB_IS_PRIMARY_KEY), "")
                tmp = tmp & IIf(prop(ENTIINFO_DB_IS_UNIQUE_KEY), ", unique_key:=" & prop(ENTIINFO_DB_IS_UNIQUE_KEY), "")
'                tmp = tmp & ", foreign_key:=""" & prop(KEY_ENTITY_FOREIGNKEY) & """"
'                tmp = tmp & ", foreign_table:=""" & prop(KEY_ENTITY_FOREIGN_TABLE) & """"
'                tmp = tmp & ", inverse_foreign_key:=""" & prop(KEY_ENTITY_INVERSE_FOREIGNKEY) & """"
'                tmp = tmp & ", inverse_foreign_table:=""" & prop(KEY_ENTITY_INVERSE_FOREIGN_TABLE) & """"
'                tmp = tmp & ", primary_key   :=""" & info(ENTIINFO_DB_IS_INDEX) & """"
                tmp = tmp & ")"
                .Add tmp
                .Add "    Call props.Add(prop, """ & prop(ENTIINFO_PROPERTY_NAME) & """)"
            Next
            .Add "    Call info.Add(props, """ & ENTIINFO_PROPERTIES & """)"
            .Add "    Call infos.Add(info, """ & info(ENTIINFO_ENTITY_NAME) & """)"
            .Add ""
            .Add ""
        Next
        .Add ""
        .Add "    Set DbInfo = infos"
        .Add "End Property"
        .Add ""
        .Add "Public Function CreateObjectBy(class_name As String) As Object"
        .Add "    Select Case class_name"
        For Each info In infos
            .Add "        Case """ & info(ENTIINFO_ENTITY_NAME) & """: Set CreateObjectBy = New " & info(ENTIINFO_ENTITY_NAME)
        Next
        .Add "    End Select"
        .Add "End Function"
        .Add ""
    End With

    For Each txt In texts: str = str + txt + vbCrLf: Next
    GetDbContextCode = str
End Function

Private Function IsIndexKeyBy(adox_indexes As ADOX.Indexes, adox_column As ADOX.Column) As Boolean
    Dim adox_index As ADOX.Index, col As ADOX.Column
    For Each adox_index In adox_indexes
        For Each col In adox_index.Columns
            If col.Name = adox_column.Name Then IsIndexKeyBy = True
        Next
    Next
End Function

Private Function IsPrimaryKeyBy(adox_indexes As ADOX.Indexes, adox_column As ADOX.Column) As Boolean
    Dim adox_index As ADOX.Index, col As ADOX.Column
    For Each adox_index In adox_indexes
        If adox_index.PrimaryKey Then
            For Each col In adox_index.Columns
                If col.Name = adox_column.Name Then IsPrimaryKeyBy = True
            Next
        End If
    Next
End Function

Private Function IsUniqueKeyBy(adox_indexes As ADOX.Indexes, adox_column As ADOX.Column) As Boolean
    Dim adox_index As ADOX.Index, col As ADOX.Column
    For Each adox_index In adox_indexes
        If adox_index.Unique Then
            For Each col In adox_index.Columns
                If col.Name = adox_column.Name Then IsUniqueKeyBy = True
            Next
        End If
    Next
End Function

Private Function GetPrimaryKeyNameByIndexes(adox_indexes As ADOX.Indexes) As String
    Dim adox_index As ADOX.Index
    For Each adox_index In adox_indexes
        If adox_index.PrimaryKey Then GetPrimaryKeyNameByIndexes = adox_index.Columns(0).Name
    Next
End Function


Private Function GetEntitesCode(entitiesName As String, entityName As String) As String
    Dim texts As New Collection, str As String, txt As Variant

    With texts
        If isExportFiles_ Then
            .Add "VERSION 1.0 CLASS"
            .Add "BEGIN"
            .Add "  MultiUse = -1  'True"
            .Add "End"
            .Add "Attribute VB_Name = """ & entitiesName & """"
            .Add "Attribute VB_GlobalNameSpace = False"
            .Add "Attribute VB_Creatable = False"
            .Add "Attribute VB_PredeclaredId = False"
            .Add "Attribute VB_Exposed = False"
        End If
        .Add "Option Explicit"
        .Add ""
        .Add "Private items_ As New Collection"
        .Add ""
        If isExportFiles_ Then
            .Add "Public Property Get NewEnum() As IUnknown"
            .Add "    Attribute NewEnum.VB_UserMemId = -4"
            .Add "    Set NewEnum = items_.[_NewEnum]"
            .Add "End Property"
        End If
        .Add "Public Property Get Item(index As Long) As " & entityName
        If isExportFiles_ Then
            .Add "Attribute item.VB_UserMemId = 0"
        End If
        .Add "    Set Item = items_.Item(index)"
        .Add "End Property"
        .Add "Public Property Get Items() As Collection"
        .Add "    Set Items = items_"
        .Add "End Property"
        .Add "Public Property Set Items(items1 As Collection)"
        .Add "    Set items_ = items1"
        .Add "End Property"
        .Add ""
        .Add "Public Function Add(T As " & entityName & ")"
        .Add "    Call items_.Add(T)"
        .Add "End Function"
        .Add "Public Function Count() As Long"
        .Add "    Count = items_.Count()"
        .Add "End Function"
        .Add "Public Function Remove(Index)"
        .Add "    Call items_.Remove(Index)"
        .Add "End Function"
        .Add ""
    End With
    For Each txt In texts: str = str + txt + vbCrLf: Next

    GetEntitesCode = str
End Function



Private Function GetEntityCode(entityName As String, props As Collection) As String
    Dim texts As New Collection, prop As Collection, str As String, txt As Variant

    With texts
        If isExportFiles_ Then
            .Add "VERSION 1.0 CLASS"
            .Add "BEGIN"
            .Add "  MultiUse = -1  'True"
            .Add "End"
            .Add "Attribute VB_Name = """ & entityName & """"
            .Add "Attribute VB_GlobalNameSpace = False"
            .Add "Attribute VB_Creatable = False"
            .Add "Attribute VB_PredeclaredId = False"
            .Add "Attribute VB_Exposed = False"
        End If
        .Add "Option Explicit"
        .Add ""
        For Each prop In props
            If IsTypeLet(prop(ENTIINFO_PROPERTY_TYPE)) Then
                .Add "Private " & prop(ENTIINFO_PROPERTY_FIELD_NAME) & " as " & prop(ENTIINFO_PROPERTY_TYPE)
            Else
                .Add "Private " & prop(ENTIINFO_PROPERTY_FIELD_NAME) & " as New " & prop(ENTIINFO_PROPERTY_TYPE)
            End If
        Next
        .Add ""
        For Each prop In props
            If IsTypeLet(prop(ENTIINFO_PROPERTY_TYPE)) Then
                .Add "Public Property Get " & prop(ENTIINFO_PROPERTY_NAME) & "() as " & prop(ENTIINFO_PROPERTY_TYPE)
                .Add "    Let " & prop(ENTIINFO_PROPERTY_NAME) & " = " & prop(ENTIINFO_PROPERTY_FIELD_NAME)
                .Add "End Property"
                .Add "Public Property Let " & prop(ENTIINFO_PROPERTY_NAME) & "(val as " & prop(ENTIINFO_PROPERTY_TYPE) & ")"
                .Add "    Let " & prop(ENTIINFO_PROPERTY_FIELD_NAME) & " = val"
                .Add "End Property"
            Else
                .Add "Public Property Get " & prop(ENTIINFO_PROPERTY_NAME) & "() as " & prop(ENTIINFO_PROPERTY_TYPE)
                .Add "    Set " & prop(ENTIINFO_PROPERTY_NAME) & " = " & prop(ENTIINFO_PROPERTY_FIELD_NAME)
                .Add "End Property"
                .Add "Public Property Set " & prop(ENTIINFO_PROPERTY_NAME) & "(val as Variant)"
                .Add "    Set " & prop(ENTIINFO_PROPERTY_FIELD_NAME) & " = val"
                .Add "End Property"
            End If
        Next
        .Add ""
    End With
    For Each txt In texts: str = str + txt + vbCrLf: Next

    GetEntityCode = str
End Function

Private Function GetPropertyByName(props As ADOX.Properties, property_name As String) As Variant
    Dim prop As ADOX.Property

    GetPropertyByName = Null
    For Each prop In props
        If prop.Name = property_name Then Set GetPropertyByName = prop
    Next

End Function


Private Function IsTypeLet(type_name As String) As Boolean
    Select Case type_name
        Case "Variant": IsTypeLet = True
        Case "Boolean": IsTypeLet = True
        Case "String": IsTypeLet = True
        Case "Date": IsTypeLet = True
        Case "Long": IsTypeLet = True
        Case "Double": IsTypeLet = True
        Case "Integer": IsTypeLet = True
        Case "Single": IsTypeLet = True
        Case "Decimal": IsTypeLet = True
        Case "Currency": IsTypeLet = True
    End Select
End Function

Private Function GetEntityPropertyTypeBy(col As ADOX.Column) As String
    Dim propertyTypeName As String, fieldType As ADODB.DataTypeEnum

    fieldType = col.Type
    If col.Attributes >= adColNullable Then
        propertyTypeName = "Variant"
    ElseIf fieldType = adInteger Then propertyTypeName = "Long"
    ElseIf fieldType = adSingle Then propertyTypeName = "Single"
    ElseIf fieldType = adDouble Then propertyTypeName = "Double"
    ElseIf fieldType = adCurrency Then propertyTypeName = "Currency"
    ElseIf fieldType = adDate Then propertyTypeName = "Date"
    ElseIf fieldType = adBoolean Then propertyTypeName = "Boolean"
    ElseIf fieldType = adDecimal Then propertyTypeName = "Decimal"
    ElseIf fieldType = adChar Then propertyTypeName = "String"
    ElseIf fieldType = adWChar Then propertyTypeName = "String"
    ElseIf fieldType = adNumeric Then propertyTypeName = "Double"
    ElseIf fieldType = adDBDate Then propertyTypeName = "Date"
    ElseIf fieldType = adDBTime Then propertyTypeName = "Date"
    ElseIf fieldType = adDBTimeStamp Then propertyTypeName = "Date"
    ElseIf fieldType = adVarNumeric Then propertyTypeName = "Double"
    ElseIf fieldType = adVarChar Then propertyTypeName = "String"
    ElseIf fieldType = adLongVarChar Then propertyTypeName = "String"
    ElseIf fieldType = adVarWChar Then propertyTypeName = "String"
    ElseIf fieldType = adLongVarWChar Then propertyTypeName = "String"
    ElseIf fieldType = adGUID Then propertyTypeName = "String"
    ElseIf fieldType = adLongVarBinary Then propertyTypeName = "Byte"
    Else
        Err.Raise 9100, "EntityContext.GetEntityPropertyTypeBy()", "field type is not supported."
    End If

    GetEntityPropertyTypeBy = propertyTypeName

End Function



Private Function GetPropertyInfo(propertyFieldName, propertyName, propertyTypeName, isLet, fieldName, relatedfieldName)
    Dim arr, prop As New Collection

    Call prop.Add(propertyName, "PropertyName")
    Call prop.Add(propertyFieldName, "ClassFieldName")
    Call prop.Add(propertyTypeName, "PropertyType")
    Call prop.Add(isLet, "IsLet")
    Call prop.Add(fieldName, "DbFieldName")
    Call prop.Add(relatedfieldName, "ForeignKey")
    Call prop.Add("", KEY_TABLE_PRIMARY_KEY)
    Set GetPropertyInfo = prop

End Function





Private Function ConvertDbTableNameToEntitesClassName(dbTbleName As String) As String
    ConvertDbTableNameToEntitesClassName = GetEstimatedPluralName(ConvertSnakeToPascalCase(ConvertToCodeChar(dbTbleName)))
End Function

Private Function ConvertDbTableNameToEntityClassName(dbTbleName As String) As String
    ConvertDbTableNameToEntityClassName = GetEstimatedSingularName(ConvertSnakeToPascalCase(ConvertToCodeChar(dbTbleName)))
End Function

Private Function ConvertDbTableNameToEntitiesClassFieldName(dbTbleName As String) As String
    ConvertDbTableNameToEntitiesClassFieldName = GetEstimatedPluralName(ConvertSnakeToCamelCase(ConvertToCodeChar(dbTbleName))) & "_"
End Function

Private Function ConvertDbFieldNameToClassPropertyName(dbFieldName As String) As String
    ConvertDbFieldNameToClassPropertyName = ConvertSnakeToPascalCase(ConvertToCodeChar(dbFieldName))
End Function

Private Function ConvertDBFieldNameToClassFieldName(dbFieldName As String) As String
    ConvertDBFieldNameToClassFieldName = ConvertSnakeToCamelCase(ConvertToCodeChar(dbFieldName)) & "_"
End Function

Private Function ConvertSnakeToPascalCase(snakeCase As String) As String
    ConvertSnakeToPascalCase = Replace(StrConv(Replace(Replace(snakeCase, "_", " "), ".", " "), vbProperCase), " ", "")
End Function

Private Function ConvertSnakeToCamelCase(snakeCase As String)
    Dim s As String
    s = Replace(StrConv(Replace(Replace(snakeCase, "_", " "), ".", " "), vbProperCase), " ", "")
    ConvertSnakeToCamelCase = StrConv(Left(s, 1), vbLowerCase) & Mid(s, 2)
End Function

Private Function GetEstimatedPluralName(singularName As String) As String
    Dim s As String

    If Right(singularName, 3) = "ses" Then s = singularName: GoTo re
    If Right(singularName, 2) = "sh" Then _
        s = Left(singularName, Len(singularName) - 2) & "shes": GoTo re
    If Right(singularName, 2) = "ch" Then _
        s = Left(singularName, Len(singularName) - 2) & "ches": GoTo re
    If Right(singularName, 1) = "y" Then _
        s = Left(singularName, Len(singularName) - 1) & "ies": GoTo re
    If Right(singularName, 1) = "x" Then _
        s = Left(singularName, Len(singularName) - 2) & "xes": GoTo re
    If Right(singularName, 1) = "s" Then s = singularName: GoTo re

    s = singularName & "s"
re:
    GetEstimatedPluralName = s
End Function

Private Function GetEstimatedSingularName(pluralName As String) As String
    Dim s As String

    If Right(pluralName, 4) = "shes" Then _
        s = Left(pluralName, Len(pluralName) - 4) & "sh": GoTo re
    If Right(pluralName, 4) = "ches" Then _
        s = Left(pluralName, Len(pluralName) - 4) & "ch": GoTo re
    If Right(pluralName, 3) = "ies" Then _
        s = Left(pluralName, Len(pluralName) - 3) & "y": GoTo re
    If Right(pluralName, 3) = "ses" Then _
        s = Left(pluralName, Len(pluralName) - 3) & "s": GoTo re
    If Right(pluralName, 3) = "xes" Then _
        s = Left(pluralName, Len(pluralName) - 3) & "x": GoTo re
    If Right(pluralName, 1) = "s" Then _
        s = Left(pluralName, Len(pluralName) - 1): GoTo re
    s = pluralName
re:
    GetEstimatedSingularName = s
End Function

Private Function ConvertToCodeChar(db_element_name As String) As String
    Dim s As String, i As Long, c As Long, ch As String
    
    db_element_name = Replace(db_element_name, "#csv", "")
    For i = 1 To Len(db_element_name)
        ch = Mid(db_element_name, i, 1)
        c = Asc(ch)
        If (48 <= c Or c <= 57 Or 65 <= c Or c <= 90 Or c = 95 Or 97 <= c Or c <= 122) Then s = s & ch
    Next i
    
    ConvertToCodeChar = s
End Function


Private Function CollectionToArray(col As Collection)
    Dim arr, c, i As Long
    ReDim arr(0 To col.Count - 1)
    i = 0
    For Each c In col
       arr(i) = c
       i = i + 1
    Next
    CollectionToArray = arr
End Function



Private Function GetEntityNameByDbTableName(dbInfos As Collection, table_name As String) As String
    Dim DbInfo As Collection
    For Each DbInfo In dbInfos
        If DbInfo(KEY_DB_TABLE_NAME) = table_name Then GetEntityNameByDbTableName = DbInfo(KEY_ENTITY_NAME): Exit Function
    Next
End Function
Private Function GetPropertyNameByDbTableFieldName(dbInfos As Collection, table_name As String, field_name As String) As String
    Dim DbInfo As Collection, prop As Collection
    For Each DbInfo In dbInfos
        If DbInfo(DBINFO_TABLE_NAME) = table_name Then
            For Each prop In DbInfo(DBINFO_PROPERTIES)
                If prop(DBINFO_FIELD_NAME) = field_name Then GetPropertyNameByDbTableFieldName = prop(DBINFO_PROPERTY_NAME): Exit Function
            Next
        End If
    Next
End Function

Private Function GetRelatedPropertyNameByDbTableFieldName(table_name As String, field_name As String)
    Dim tbl As ADOX.Table, ky As ADOX.Key, col As ADOX.Column
    
    For Each ky In catalog_.Tables(table_name).Keys
        If ky.Type = adKeyForeign Then
            For Each col In ky.Columns
                If col.Name = field_name Then GetRelatedPropertyNameByDbTableFieldName = GetPropertyNameByDbTableFieldName(dbInfos, tble_name, field_name)
            Next
        End If
    Next
    
    
End Function


Private Function FetchRecordSetByQuery(query_string As String) As ADODB.Recordset
    OpenDb
    Set FetchRecordSetByQuery = New ADODB.Recordset
    FetchRecordSetByQuery.CursorLocation = adUseClient
    Call FetchRecordSetByQuery.Open(query_string, cn_, LockType:=adLockReadOnly, CursorType:=adOpenStatic)
End Function


' Get value from DB informaiton --------------------------------------------------------------------------------------------------------------------
Private Function GetFieldNameFromDbInfoBy(db_info, entity_name As String, Optional prop_name As String)
    If prop_name <> "" Then GetFieldNameFromDbInfoBy = db_info(entity_name)(DBINFO_PROPERTIES)(prop_name)(DBINFO_FIELD_NAME)
End Function

Private Function GetEntityNameFromDbInfoBy(db_infos, Optional entities_name As String)
    Dim dinfo, prop
    
    If entities_name <> "" Then
        For Each dinfo In db_infos
            If dinfo(DBINFO_ENTITIES_NAME) = entities_name Then GetEntityNameFromDbInfoBy = dinfo(DBINFO_ENTITY_NAME): Exit Function
        Next
    End If
End Function

Private Function SetFieldNamesStringFromDbInfoBy(db_infos, entity_name1 As String, entity_name2 As String) As String
    Dim prop, fn As String, tbl_name As String
    
    For Each prop In db_infos(entity_name1)(DBINFO_PROPERTIES)
        fn = prop(DBINFO_FIELD_NAME)
        If fn <> "" And fetchFieldNames_.Exists(fn) = False Then Call fetchFieldNames_.Add(fn, fn)
    Next
    For Each prop In db_infos(entity_name2)(DBINFO_PROPERTIES)
        fn = prop(DBINFO_FIELD_NAME)
        If fn <> "" And fetchFieldNames_.Exists(fn) = False Then Call fetchFieldNames_.Add(fn, fn)
    Next
    
End Function

Private Function GetFieldNamesStringFromDbInfoBy(db_infos, entity_name As String) As String
    Dim fns As New Dictionary, prop, fn As String, tbl_name As String
    
    For Each prop In db_infos(entity_name)(DBINFO_PROPERTIES)
        fn = prop(DBINFO_FIELD_NAME)
        If fn <> "" And fns.Exists(fn) = False Then Call fns.Add(fn, fn)
    Next
    
    GetFieldNamesStringFromDbInfoBy = Join(fns.Keys, ",")
End Function





' Conditions -------------------------------------------------------------------------
Public Function CanGetAdoxKey() As Boolean
   CanGetAdoxKey = InStr(cn_.PROVIDER, "SQL") = 0
End Function



' Useful functions -------------------------------------------------------------------
Private Function TrimLast(txt, trim_text) As String
    TrimLast = Left(txt, Len(txt) - Len(trim_text))
End Function

Private Function CollectionContain(col As Collection, val) As Boolean
    Dim c
    For Each c In col
        If c = val Then CollectionContain = True: Exit Function
    Next
End Function

