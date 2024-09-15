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
' Version   : 0.5.0
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
Private Const DBINFO_PROPERTY_TYPE_NAME As String = "PropertyTypeName"
Private Const DBINFO_FIELD_NAME As String = "DbFieldName"
Private Const DBINFO_FIELD_TYPE As String = "DbFieldType"
Private Const DBINFO_IS_AUTONUMBER As String = "DbIsAutoNumber"
Private Const DBINFO_IS_PRIMARY_KEY As String = "DbIsPrimaryKey"
Private Const DBINFO_IS_UNIQUE_KEY As String = "DbIsUniqueKey"
Private Const DBINFO_INDEX As String = "DbIsPrimaryKey"
Private Const DBINFO_BASE_PROPERTY_NAME As String = "DbBaseProperty"
Private Const DBINFO_RELATED_PROPERTY_NAME As String = "DbRelatedProperty"

Private Const INCLUDE_ENTITY_NAME As String = "EntityName"
Private Const INCLUDE_PROPERTIES As String = "Properties"


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
Private isTransactionActive_ As Boolean

Private baseEntityName_ As String
Private includedEntityNames_ As Collection
Private lastIncludedEntityProps_ As Collection
Private fetchFieldNames_ As New Dictionary

Private querySelect_ As String
Private queryField_ As String
Private queryFrom_ As String
Private queryJoin_ As String
Private queryJoins_ As Dictionary
Private queryWhere_ As String
Private queryOrder_ As String
Private lastIncludedEntityName_ As String
Private includedConnectTableString_ As String
Private queryFields_ As New Dictionary

Private dbInfo_ As Collection


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
    DisconnectDb
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else Set cn_ = Nothing
End Sub

' Properties -------------------------------------------------------------------------------
Public Property Set DbInfo(v As Object)
    Set dbInfo_ = v
End Property


' Methods ----------------------------------------------------------------------------------
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
    Optional property_type As String, _
    Optional db_field_name As String, _
    Optional db_field_type As ADOX.DataTypeEnum = 0, _
    Optional auto_increment As Boolean = False, _
    Optional primary_key As Boolean = False, _
    Optional unique_key As Boolean = False, _
    Optional base_property_name As Variant, _
    Optional related_property_name As Variant _
    )
    Dim prop As Collection
    
    Set prop = New Collection
    Call prop.Add(property_name, DBINFO_PROPERTY_NAME)
    Call prop.Add(property_type, DBINFO_PROPERTY_TYPE_NAME)
    Call prop.Add(db_field_name, DBINFO_FIELD_NAME)
    Call prop.Add(db_field_type, DBINFO_FIELD_TYPE)
    Call prop.Add(auto_increment, DBINFO_IS_AUTONUMBER)
    Call prop.Add(primary_key, DBINFO_IS_PRIMARY_KEY)
    Call prop.Add(unique_key, DBINFO_IS_UNIQUE_KEY)
    
    ' todo apply multi relationship by using array
    Call prop.Add(base_property_name, DBINFO_BASE_PROPERTY_NAME)
    Call prop.Add(related_property_name, DBINFO_RELATED_PROPERTY_NAME)
    
    Set CreateDbInfoPropertyBy = prop
End Function

Public Function SetQueryDbSet(entity_name) As String
    baseEntityName_ = entity_name
    Set queryFields_ = New Dictionary

    includedConnectTableString_ = ""
    includedConnectTableString_ = baseEntityName_
    
    querySelect_ = "select "
    queryField_ = ""
    queryJoin_ = ""
    Set queryJoins_ = New Dictionary
    queryWhere_ = ""
    queryOrder_ = ""
    
    Call SetQueryField((dbInfo_(baseEntityName_)(DBINFO_SQL_TABLE_NAME)), dbInfo_(baseEntityName_)(DBINFO_PROPERTIES))
    queryFrom_ = "from " & dbInfo_(baseEntityName_)(DBINFO_SQL_TABLE_NAME) & vbCrLf
End Function

Public Function SetQuerySelect(db_info, ParamArray property_names())
End Function

Public Function SetQueryWhere(property_name As String, sql_operator As String, compare_value As String, Optional entity_name As String)
    entity_name = IIf(entity_name = "", baseEntityName_, entity_name)
    If queryWhere_ = "" Then
        queryWhere_ = " where " & GetDbInfoSqlFieldNameBy(entity_name, property_name) & " " & sql_operator & " " & compare_value
    Else
        queryWhere_ = queryWhere_ & " and " & GetDbInfoSqlFieldNameBy(entity_name, property_name) & " " & sql_operator & " " & compare_value
    End If
End Function

Public Function SetQueryOrderBy(property_name As String, Optional entity_name As String, Optional is_asc As Boolean) As String
    entity_name = IIf(entity_name = "", baseEntityName_, entity_name)
    queryOrder_ = " order by " & (GetDbInfoSqlFieldNameBy(entity_name, property_name)) & IIf(is_asc, " asc", " desc")
End Function

Public Function SetQueryThenBy(property_name As String, Optional entity_name As String, Optional is_asc As Boolean) As String
    entity_name = IIf(entity_name = "", baseEntityName_, entity_name)
    queryOrder_ = queryOrder_ & " , " & (GetDbInfoSqlFieldNameBy(entity_name, property_name)) & IIf(is_asc, " asc", " desc")
End Function

Public Function SetQueryInclude(property_name As String) As String
    Dim base_table As String, base_field As String, related_table As String, related_field As String, base_prop_name As String, related_prop_name As String, related_entity_name As String
    Dim prop As Collection, full_name As String, qry As String
        
    base_prop_name = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_BASE_PROPERTY_NAME)
    related_prop_name = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_RELATED_PROPERTY_NAME)
    related_entity_name = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_PROPERTY_TYPE_NAME)
    base_table = dbInfo_(baseEntityName_)(DBINFO_SQL_TABLE_NAME)
    base_field = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(base_prop_name)(DBINFO_FIELD_NAME)
    related_table = dbInfo_(related_entity_name)(DBINFO_SQL_TABLE_NAME)
    related_field = dbInfo_(related_entity_name)(DBINFO_PROPERTIES)(related_prop_name)(DBINFO_FIELD_NAME)
    
    qry = " left outer join " & related_table & " on " & base_table & ".[" & base_field & "] = " & related_table & ".[" & related_field & "]" & vbCrLf
    If queryJoins_.Exists(qry) = False Then
        queryJoin_ = queryJoin_ & qry
        Call queryJoins_.Add(qry, qry)
    End If
    
    includedConnectTableString_ = baseEntityName_
    includedConnectTableString_ = includedConnectTableString_ & property_name
    Call SetQueryField(related_table, dbInfo_(related_entity_name)(DBINFO_PROPERTIES))
       
    lastIncludedEntityName_ = related_entity_name
    
End Function

Public Function SetQueryThenInclude(property_name As String) As String
    Dim base_table As String, base_field As String, related_table As String, related_field As String, base_prop_name As String, related_prop_name As String, related_entity_name As String
    Dim prop As Collection, full_name As String, qry As String
    
    base_prop_name = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_BASE_PROPERTY_NAME)
    related_prop_name = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_RELATED_PROPERTY_NAME)
    related_entity_name = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_PROPERTY_TYPE_NAME)
    base_table = dbInfo_(lastIncludedEntityName_)(DBINFO_SQL_TABLE_NAME)
    base_field = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(base_prop_name)(DBINFO_FIELD_NAME)
    related_table = dbInfo_(related_entity_name)(DBINFO_SQL_TABLE_NAME)
    related_field = dbInfo_(related_entity_name)(DBINFO_PROPERTIES)(related_prop_name)(DBINFO_FIELD_NAME)
    
    qry = " left outer join " & related_table & " on " & base_table & ".[" & base_field & "] = " & related_table & ".[" & related_field & "]" & vbCrLf
    If queryJoins_.Exists(qry) = False Then
        queryJoin_ = queryJoin_ & qry
        Call queryJoins_.Add(qry, qry)
    End If
    
    includedConnectTableString_ = includedConnectTableString_ & property_name

    Call SetQueryField(related_table, dbInfo_(related_entity_name)(DBINFO_PROPERTIES))
      
    lastIncludedEntityName_ = related_entity_name
End Function

Public Function FetchRecords(dbc As Object)
    Dim rs As Recordset, qry As String, re As Collection, records_rows As Collection
    
    qry = querySelect_ & vbCrLf & queryField_ & vbCrLf & queryFrom_ & vbCrLf & queryJoin_ & vbCrLf & queryWhere_ & vbCrLf & queryOrder_
    
    Call OpenDb
    Set rs = FetchRecordSetByQuery(qry)
    If rs.BOF And rs.EOF Then GoTo errFetchQuery
        
    Set records_rows = ConvertRecordRowsToCollection(rs.GetRows)
    Set re = FetchRecsSubRecursive(records_rows, baseEntityName_, baseEntityName_, dbc, cnt:=0)
    
    ' todo take, skip, first
    Set FetchRecords = re
    Exit Function
        
errFetchQuery:
    Set FetchRecords = Nothing:     Exit Function
    
End Function

Public Function SaveChanges()
On Error GoTo errCommit
    Call CommitTransaction
    Exit Function
errCommit:
    Call RollbackTransaction
    Call Err.Raise(100000, Err.Source & vbCrLf & "EntityContexter.SaveChanges()", "Failed to save changes")
End Function



Public Function Add(rec As Object)
    Dim rs As New ADODB.Recordset, prop As Collection
    Dim ent_name, is_related_field_index As Boolean, r As Object, v As Variant
                                            
On Error GoTo errOpen
    OpenDb
    cn_.CursorLocation = adUseClient
    
    Call AddRelatedEntity(rec, dbInfo_)
                
    If HasPrimaryKey(rec, dbInfo_) = False Then If ExistUniqueRecord(rec, dbInfo_) Then Call Err.Raise(10003, "EntityContexter.Add()", "Record is already exist")
    Call BeginTransaction
    Call rs.Open(Source:=dbInfo_(TypeName(rec))(KEY_SQL_TABLE_NAME), ActiveConnection:=cn_, LockType:=adLockOptimistic)
    Call rs.AddNew(GetFieldsListByEntity(rec, dbInfo_), GetValuesListByEntity(rec, dbInfo_))
    Call SetAutoincrementFieldsValues(rs, rec, dbInfo_)
    
    Call AddRelatedEntities(rec, dbInfo_)
    Exit Function

errOpen:
    DisconnectDb
    Call Err.Raise(Err.Number, Err.Source & vbCrLf & "EntityContexter.Add()", Err.Description)
End Function

Private Function AddRelatedEntity(rec As Object, DbInfo As Object)
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean, r As Object, v As Variant
    Dim related_entity_properties As New Collection
    
    Set related_entity_properties = GetRelatedEntityEntitiesProperty(rec, DbInfo, is_entity:=True)
    
    For Each prop In related_entity_properties
        Set r = CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet)
        Dim pky_value: pky_value = CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbGet)
        If IsNullOrEmptyOrZeroOrWhiteSpace(pky_value) Then
            Call Add(r)
            v = CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbGet)
            Call CallByName(rec, prop(DBINFO_BASE_PROPERTY_NAME), VbLet, (v))
        End If
    Next
End Function
Private Function AddRelatedEntities(rec As Object, DbInfo As Object)
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean, r As Object, v As Variant
    Dim related_entities_properties As New Collection

    Set related_entities_properties = GetRelatedEntityEntitiesProperty(rec, DbInfo, is_entity:=False)
    
    For Each prop In related_entities_properties
        For Each r In CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet)
            v = CallByName(rec, prop(DBINFO_BASE_PROPERTY_NAME), VbGet)
            Call CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbLet, (v))
            Call Add(r)
        Next
    Next
    
End Function
Private Function GetRelatedEntityEntitiesProperty(rec As Object, DbInfo As Object, is_entity As Boolean) As Collection
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean
    Dim related_properties As New Collection

    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) = "" Then
            ent_name = GetEntityName(rec, prop, DbInfo)
            If ent_name <> "Nothing" Then
                is_related_field_index = DbInfo(ent_name)(DBINFO_PROPERTIES)(prop(DBINFO_RELATED_PROPERTY_NAME))(DBINFO_IS_UNIQUE_KEY)
                If is_related_field_index = is_entity Then
                    Call related_properties.Add(prop)
                End If
            End If
        End If
    Next
    Set GetRelatedEntityEntitiesProperty = related_properties
End Function
Private Function GetEntityName(ent As Object, prop As Collection, db_info) As String
    Dim p, ent_name
    Set p = CallByName(ent, prop(DBINFO_PROPERTY_NAME), VbGet)
    ent_name = GetEntityNameFromDbInfoBy(db_info, TypeName(p))
    If ent_name <> "" Then Else ent_name = TypeName(p)
    GetEntityName = ent_name
End Function



Private Function GetValuesListByEntity(rec As Object, DbInfo As Object)
    Dim prop As Collection, col As New Collection
    
    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" And prop(DBINFO_IS_AUTONUMBER) = False Then Call col.Add(CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet))
    Next
    GetValuesListByEntity = CollectionToArray(col)

End Function

Private Function GetFieldsListByEntity(rec As Object, DbInfo As Object)
    Dim prop As Collection, col As New Collection

    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" And prop(DBINFO_IS_AUTONUMBER) = False Then Call col.Add(prop(DBINFO_FIELD_NAME))
    Next
    GetFieldsListByEntity = CollectionToArray(col)

End Function

Public Function Update(rec As Object)
    
On Error GoTo errOpen
    Call OpenDb
    cn_.CursorLocation = adUseClient
    
    If HasPrimaryKey(rec, dbInfo_) = False Then Call Err.Raise(10001, "EntityContexter.Update()", "Table has no primary key")
    If ExistUniqueRecord(rec, dbInfo_) = False Then Call Err.Raise(10002, "EntityContexter.Update()", "Not found records")
    
    Call BeginTransaction
    Call cn_.Execute(CommandText:=CreateUpdateQuery(rec, dbInfo_))
    Exit Function

errOpen:
    DisconnectDb
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Update()", Err.Description
End Function

Public Function Remove(rec As Object)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    cn_.CursorLocation = adUseClient
    If HasPrimaryKey(rec, dbInfo_) = False Then Call Err.Raise(10001, "EntityContexter.Remove()", "Table has no primary key")
    If ExistUniqueRecord(rec, dbInfo_) = False Then Call Err.Raise(10002, "EntityContexter.Remove()", "Not found records")
    
    Call BeginTransaction
    Call cn_.Execute(CommandText:=CreateDeleteQuery(rec, dbInfo_))
    Exit Function

errOpen:
    DisconnectDb
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

    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_IS_AUTONUMBER) = True Then _
            Call CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbLet, rs.Fields(prop(DBINFO_FIELD_NAME)).Value)
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

Private Function GetQueryPrimaryWhere(rec As Object, db_Infos As Object) As String
    Dim qry As String, prop As Collection
    
    qry = " Where "
    For Each prop In db_Infos(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_IS_PRIMARY_KEY) Then
            qry = qry & db_Infos(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & ".[" & prop(DBINFO_FIELD_NAME) & "] = " & ConvertValueToSqlQuery(prop(DBINFO_PROPERTY_NAME), rec, db_Infos) & ", "
        End If
    Next
    qry = Left(qry, Len(qry) - 2)
    
    GetQueryPrimaryWhere = qry
End Function

Private Function GetPropByTableFeildName(db_Infos As Collection, table_name As String, field_name As String) As Collection
    Dim db_info, props, prop
    For Each db_info In db_Infos
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
        If prop(DBINFO_FIELD_NAME) <> "" Then
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
    Dim rel, rels
    
    Set rels = GetRelations1(base_table.Name)
    
    ' todo
    For Each tbl In catalog_.Tables
        Set rels = GetRelations(base_table.Name, tbl.Name)
        For Each rel In rels
            If rel(KEY_DB_BASE_KEY_TABLE_NAME) = base_table.Name And rel(KEY_DB_BASE_KEY_FIELD_NAME) = base_column.Name Then
                If rel("IsNto1") Then
                    Set prop = GetPropertyInfoOfRelatedEntity((rel(KEY_DB_RELATED_TABLE_NAME)))
                Else
                    Set prop = GetPropertyInfoOfRelatedEntities((rel(KEY_DB_RELATED_TABLE_NAME)))
                End If
                Call props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
            End If
        Next
    Next

End Function


Private Function GetRelations1(base_table_name As String) As Collection
    Dim tbl As ADOX.Table, ky As ADOX.Key, col As ADOX.Column, key_infos As New Collection, key_info As New Collection
    Call OpenDb
    
    If CanGetAdoxKey Then
        For Each ky In catalog_.Tables(base_table_name).Keys
            If ky.Type = adKeyForeign Then
                For Each col In ky.Columns
                    Set key_info = New Collection
                    Call key_info.Add(base_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
                    Call key_info.Add(ky.RelatedTable, KEY_DB_RELATED_TABLE_NAME)
                    Call key_info.Add(col.Name, KEY_DB_BASE_KEY_FIELD_NAME)
                    Call key_info.Add(col.RelatedColumn, KEY_DB_RELATED_FIELD_NAME)
                    Call key_info.Add(True, "IsNto1")
                    Call key_infos.Add(key_info, ky.Name)
                Next
            End If
        Next
        For Each tbl In catalog_.Tables
            For Each ky In tbl.Keys
                If ky.RelatedTable = base_table_name And ky.Type = adKeyForeign Then
                    For Each col In ky.Columns
                        Set key_info = New Collection
                        Call key_info.Add(base_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
                        Call key_info.Add(tbl.Name, KEY_DB_RELATED_TABLE_NAME)
                        Call key_info.Add(col.RelatedColumn, KEY_DB_BASE_KEY_FIELD_NAME)
                        Call key_info.Add(col.Name, KEY_DB_RELATED_FIELD_NAME)
                        Call key_info.Add(False, "IsNto1")
                        Call key_infos.Add(key_info, ky.Name)
                    Next
                 End If
            Next
        Next
    Else
        Dim qry As String, db_name As String
        db_name = "[" & cn_.DefaultDatabase & "]"
        qry = "SELECT fk.CONSTRAINT_NAME AS ForeignKey, fk.TABLE_NAME AS ChildTable, fk.COLUMN_NAME AS ChildColumn, pk.TABLE_NAME AS ParentTable, pk.COLUMN_NAME AS ParentColumn " & _
            "FROM " & db_name & ".[INFORMATION_SCHEMA].REFERENTIAL_CONSTRAINTS AS rc " & _
            "JOIN " & db_name & ".[INFORMATION_SCHEMA].KEY_COLUMN_USAGE AS fk ON rc.CONSTRAINT_NAME = fk.CONSTRAINT_NAME " & _
            "JOIN " & db_name & ".[INFORMATION_SCHEMA].KEY_COLUMN_USAGE AS pk ON rc.UNIQUE_CONSTRAINT_NAME = pk.CONSTRAINT_NAME Order BY ChildTable, ForeignKey;"
        
        Dim rs, i As Long
        rs = cn_.Execute(qry).GetRows
        For i = LBound(rs, 2) To UBound(rs, 2)
            Set key_info = New Collection
            If rs(1, i) = base_table_name Then
                Call key_info.Add(rs(1, i), KEY_DB_BASE_KEY_TABLE_NAME)
                Call key_info.Add(rs(3, i), KEY_DB_RELATED_TABLE_NAME)
                Call key_info.Add(rs(2, i), KEY_DB_BASE_KEY_FIELD_NAME)
                Call key_info.Add(rs(4, i), KEY_DB_RELATED_FIELD_NAME)
                Call key_info.Add(True, "IsNto1")
                Call key_infos.Add(key_info, rs(0, i))
            ElseIf rs(3, i) = base_table_name Then
                Call key_info.Add(rs(3, i), KEY_DB_BASE_KEY_TABLE_NAME)
                Call key_info.Add(rs(1, i), KEY_DB_RELATED_TABLE_NAME)
                Call key_info.Add(rs(4, i), KEY_DB_BASE_KEY_FIELD_NAME)
                Call key_info.Add(rs(2, i), KEY_DB_RELATED_FIELD_NAME)
                Call key_info.Add(False, "IsNto1")
                Call key_infos.Add(key_info, rs(0, i))
            End If
        Next
        
    End If
    
    Set GetRelations1 = key_infos
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
                    Call key_info.Add(True, "IsNto1")
                    Call key_infos.Add(key_info, ky.Name)
                Next
            End If
        Next
        For Each tbl In catalog_.Tables
            If tbl.Name = related_table_name Then
                For Each ky In tbl.Keys
                    If ky.RelatedTable = base_table_name And ky.Type = adKeyForeign Then
                        For Each col In ky.Columns
                            Call key_info.Add(base_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
                            Call key_info.Add(related_table_name, KEY_DB_RELATED_TABLE_NAME)
                            Call key_info.Add(col.RelatedColumn, KEY_DB_BASE_KEY_FIELD_NAME)
                            Call key_info.Add(col.Name, KEY_DB_RELATED_FIELD_NAME)
                            Call key_info.Add(False, "IsNto1")
                            Call key_infos.Add(key_info, ky.Name)
                        Next
                     End If
                Next
            End If
        Next
    Else
        Dim qry As String
        Dim db_name As String
        db_name = cn_.DefaultDatabase
        qry = "SELECT fk.CONSTRAINT_NAME AS ForeignKey, fk.TABLE_NAME AS ChildTable, fk.COLUMN_NAME AS ChildColumn, pk.TABLE_NAME AS ParentTable, pk.COLUMN_NAME AS ParentColumn " & _
            "FROM " & db_name & ".[INFORMATION_SCHEMA].REFERENTIAL_CONSTRAINTS AS rc " & _
            "JOIN " & db_name & ".[INFORMATION_SCHEMA].KEY_COLUMN_USAGE AS fk ON rc.CONSTRAINT_NAME = fk.CONSTRAINT_NAME " & _
            "JOIN " & db_name & ".[INFORMATION_SCHEMA].KEY_COLUMN_USAGE AS pk ON rc.UNIQUE_CONSTRAINT_NAME = pk.CONSTRAINT_NAME Order BY ChildTable, ForeignKey;"
        
        Dim rs, i As Long
        rs = cn_.Execute(qry).GetRows
        For i = LBound(rs, 2) To UBound(rs, 2)
            If rs(1, i) = base_table_name And rs(3, i) = related_table_name Then
                Call key_info.Add(rs(1, i), KEY_DB_BASE_KEY_TABLE_NAME)
                Call key_info.Add(rs(3, i), KEY_DB_RELATED_TABLE_NAME)
                Call key_info.Add(rs(2, i), KEY_DB_BASE_KEY_FIELD_NAME)
                Call key_info.Add(rs(4, i), KEY_DB_RELATED_FIELD_NAME)
                Call key_info.Add(True, "IsNto1")
                Call key_infos.Add(key_info, rs(0, i))
            ElseIf rs(1, i) = related_table_name And rs(3, i) = base_table_name Then
                Call key_info.Add(rs(3, i), KEY_DB_BASE_KEY_TABLE_NAME)
                Call key_info.Add(rs(1, i), KEY_DB_RELATED_TABLE_NAME)
                Call key_info.Add(rs(4, i), KEY_DB_BASE_KEY_FIELD_NAME)
                Call key_info.Add(rs(2, i), KEY_DB_RELATED_FIELD_NAME)
                Call key_info.Add(False, "IsNto1")
                Call key_infos.Add(key_info, rs(0, i))
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
        .Add "        Call ec_.SetQueryDbSet(entity_name):     Set DbSet = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function Where(property_name As String, sql_operator As String, compare_value As String, Optional entity_name As String) As " & dbc_name
        .Add "        Call ec_.SetQueryWhere(property_name, sql_operator, compare_value, entity_name):      Set Where = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function Include(entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryInclude(entity_name_, entity_name):     Set Include = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function OrderBy(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryOrderBy(property_name, entity_name, true):   Set OrderBy = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function OrderByDescending(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryOrderBy(property_name, entity_name, false):   Set OrderByDescending = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenBy(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryThenBy(property_name, entity_name, false):   Set ThenBy = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenByDescending(property_name As String, Optional entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryThenBy(property_name, entity_name, false):   Set ThenByDescending = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ToList() As Variant"
        .Add "    Set ToList = ec_.FetchRecords(Me)"
        .Add "End Function"
        .Add ""
        .Add ""
        .Add "Public Function Add(rec As Object)"
        .Add ""
        .Add "    Call ec_.Add(rec)"
        .Add "    Call ec_.SaveChanges(rec)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Update(rec As Object)"
        .Add ""
        .Add "    Call ec_.Update(rec)"
        .Add "    Call ec_.SaveChanges(rec)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Remove(rec As Object)"
        .Add ""
        .Add "    Call ec_.Remove(rec)"
        .Add "    Call ec_.SaveChanges(rec)"
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
                tmp = tmp & IIf(prop(ENTIINFO_DB_IS_AUTOINCREMENT), ", auto_increment:=" & prop(ENTIINFO_DB_IS_AUTOINCREMENT), "")
                tmp = tmp & IIf(prop(ENTIINFO_DB_IS_PRIMARY_KEY), ", primary_key:=" & prop(ENTIINFO_DB_IS_PRIMARY_KEY), "")
                tmp = tmp & IIf(prop(ENTIINFO_DB_IS_UNIQUE_KEY), ", unique_key:=" & prop(ENTIINFO_DB_IS_UNIQUE_KEY), "")
                ' todo add relatetion
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
            .Add "        Case """ & info(ENTIINFO_ENTITIES_NAME) & """: Set CreateObjectBy = New " & info(ENTIINFO_ENTITIES_NAME)
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
            .Add "Private " & prop(ENTIINFO_PROPERTY_FIELD_NAME) & " as " & prop(ENTIINFO_PROPERTY_TYPE)
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


' CreateQuery() funcitons --------------------------------------------------------------------------------------------------------------------------
Private Function SetQueryField(table_name As String, props As Variant)
    Dim prop As Collection, full_name As String
            
    If queryField_ <> "" Then queryField_ = queryField_ & ", "
    For Each prop In props
        If prop(DBINFO_FIELD_NAME) <> "" Then
            full_name = includedConnectTableString_ & prop(DBINFO_PROPERTY_NAME)
            If queryFields_.Exists(full_name) = False Then
                queryField_ = queryField_ & vbCrLf & table_name & ".[" & prop(DBINFO_FIELD_NAME) & "]" & " as " & ConvertDbInfoFullNameToColumnName(full_name) & ", "
                Call queryFields_.Add(full_name, queryFields_.Count + 1)
            End If
        End If
    Next
    
    queryField_ = Left(queryField_, Len(queryField_) - Len(", "))
End Function

Private Function ConvertDbInfoFullNameToColumnName(full_field_name As String) As String
    ConvertDbInfoFullNameToColumnName = Replace(Replace(Replace(full_field_name, ".", ""), "[", ""), "]", "")
End Function


' FetchRecords() functions ------------------------------------------------------------------------------------------------------------------------
Private Function FetchRecsSubRecursive(records_row As Collection, ByVal entity_name As String, ByVal prop_path As String, dbc, Optional cnt As Long) As Object
    Dim recs As New Collection, grouped_recs As Collection, record_rows As Collection, record_row, rec As Object, field_name As String, prop_name, v As Variant
    Dim prop As Collection, re As Collection, entities_object As Object
    
    If ExistPropertyPath(entity_name, prop_path) = False Then Set FetchRecsSubRecursive = New Collection: Exit Function
    
    Set grouped_recs = GroupRecordsCollectionByEntityName(records_row, entity_name, prop_path)
    For Each record_rows In grouped_recs
        If record_rows.Count > 0 Then record_row = record_rows(1)
        If IsRecordAllNullValue(record_row, entity_name, prop_path) = False Then
            ' to create object
            Set rec = dbc.CreateObjectBy(entity_name)
            Call recs.Add(rec)
            
            ' set values
            For Each prop_name In GetPropertyNamesFromDbInfoBy(entity_name)
                field_name = prop_path & prop_name
                If queryFields_.Exists(field_name) Then
                    v = record_row(queryFields_.Item(field_name))
                    Call CallByName(rec, prop_name, VbLet, (v))
                End If
            Next
                
            ' to recul related object
            For Each prop In GetRelatedPropFromDbInfoBy(entity_name)
                prop_name = prop(DBINFO_PROPERTY_NAME)
                Set re = FetchRecsSubRecursive(record_rows, prop(DBINFO_PROPERTY_TYPE_NAME), prop_path & prop_name, dbc)
                
                If re.Count = 0 Then
                    Call CallByName(rec, prop_name, VbSet, Nothing)
                Else
                    If IsRelatedPropertyUniqueKey(prop) Then
                        Call CallByName(rec, prop_name, VbSet, re(1))
                    Else
                        Set entities_object = dbc.CreateObjectBy((dbInfo_(prop(DBINFO_PROPERTY_TYPE_NAME))(DBINFO_ENTITIES_NAME)))
                        Set entities_object.Items = re
                        Call CallByName(rec, prop_name, VbSet, entities_object)
                    End If
                End If
            Next
        End If
    Next
        
    Set FetchRecsSubRecursive = recs
End Function

Private Function GroupRecordsCollectionByEntityName(records_collection As Collection, entity_name As String, prop_path As String) As Collection
    Dim r As Variant, i As Long, records_with_index As Dictionary, re As New Collection, same_entity_records As New Collection
    
    Set records_with_index = ConvertCollectionToDictionaryWithIndexKey(records_collection)
    
    For i = 1 To records_with_index.Count
        r = records_with_index.Item(CStr(i))
        If HasSameRecord(r, same_entity_records, entity_name, prop_path) Then
            Call same_entity_records.Add((r))
            Call records_with_index.Remove(CStr(i))
        Else
            Set same_entity_records = New Collection
            Call re.Add(same_entity_records)
            Call same_entity_records.Add(r)
            Call records_with_index.Remove(CStr(i))
        End If
    Next
    
    Set GroupRecordsCollectionByEntityName = re
End Function

Private Function HasSameRecord(r, col As Collection, entity_name As String, prop_path) As Boolean
    Dim prop_name, field_name As String, c As Long, r1, f As Boolean
    
    For Each r1 In col
        f = True
        For Each prop_name In GetPropertyNamesFromDbInfoBy(entity_name)
            field_name = prop_path & prop_name
            c = queryFields_(field_name)
            If r(c) <> r1(c) Then f = False
        Next
        If f Then HasSameRecord = True: Exit Function
    Next

End Function

Private Function ExistPropertyPath(entity_name As String, prop_path As String) As Boolean
    Dim prop_name, prop As Collection
    
    For Each prop In dbInfo_(entity_name)(DBINFO_PROPERTIES)
        prop_name = prop_path & prop(DBINFO_PROPERTY_NAME)
        If queryFields_.Exists(prop_name) Then ExistPropertyPath = True: Exit Function
    Next

End Function

Private Function IsRecordAllNullValue(record_array, entity_name As String, prop_path As String) As Boolean
    Dim prop_name, field_name
    
    IsRecordAllNullValue = True
    For Each prop_name In GetPropertyNamesFromDbInfoBy(entity_name)
        field_name = prop_path & prop_name
        If queryFields_.Exists(field_name) Then If IsNull(record_array(queryFields_.Item(field_name))) = False Then IsRecordAllNullValue = False: Exit Function
    Next

End Function


' connections -------------------------------------------------------------------------------------------------------------------------------------
Private Function BeginTransaction()
    If cn_.State <> adStateClosed And isTransactionActive_ = False Then
        Call cn_.BeginTrans
        isTransactionActive_ = True
    End If
End Function

Private Function CommitTransaction()
    If cn_.State <> adStateClosed And isTransactionActive_ Then
        Call cn_.CommitTrans
        isTransactionActive_ = False
    End If
End Function

Private Function RollbackTransaction()
    If cn_.State <> adStateClosed And isTransactionActive_ Then
        Call cn_.RollbackTrans
        isTransactionActive_ = False
    End If
End Function

Private Function DisconnectDb()
    Call RollbackTransaction
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Function



' Get value from DB informaiton --------------------------------------------------------------------------------------------------------------------
Private Function GetFieldNameFromDbInfoBy(db_info, entity_name As String, Optional prop_name As String)
    If prop_name <> "" Then GetFieldNameFromDbInfoBy = db_info(entity_name)(DBINFO_PROPERTIES)(prop_name)(DBINFO_FIELD_NAME)
End Function

Private Function GetEntityNameFromDbInfoBy(db_Infos, Optional entities_name As String)
    Dim dinfo, prop
    
    If entities_name <> "" Then
        For Each dinfo In db_Infos
            If dinfo(DBINFO_ENTITIES_NAME) = entities_name Then GetEntityNameFromDbInfoBy = dinfo(DBINFO_ENTITY_NAME): Exit Function
        Next
    End If
End Function

Private Function SetFieldNamesStringFromDbInfoBy(db_Infos, entity_name1 As String, entity_name2 As String) As String
    Dim prop, fn As String, tbl_name As String
    
    For Each prop In db_Infos(entity_name1)(DBINFO_PROPERTIES)
        fn = prop(DBINFO_FIELD_NAME)
        If fn <> "" And fetchFieldNames_.Exists(fn) = False Then Call fetchFieldNames_.Add(fn, fn)
    Next
    For Each prop In db_Infos(entity_name2)(DBINFO_PROPERTIES)
        fn = prop(DBINFO_FIELD_NAME)
        If fn <> "" And fetchFieldNames_.Exists(fn) = False Then Call fetchFieldNames_.Add(fn, fn)
    Next
    
End Function

Private Function GetFieldNamesStringFromDbInfoBy(db_Infos, entity_name As String, Optional use_table_name As Boolean = True) As String
    Dim fns As New Dictionary, prop, fn As String, tbl_name As String
    
    For Each prop In db_Infos(entity_name)(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" Then
            If use_table_name Then fn = db_Infos(entity_name)(DBINFO_SQL_TABLE_NAME) & ".[" & prop(DBINFO_FIELD_NAME) & "]" Else fn = prop(DBINFO_FIELD_NAME)
            If fn <> "" And fns.Exists(fn) = False Then Call fns.Add(fn, fn)
        End If
    Next
    
    GetFieldNamesStringFromDbInfoBy = Join(fns.Keys, ",")
End Function

Private Function GetPropertyNamesFromDbInfoBy(entity_name As String) As Collection
    Dim col As New Collection, prop As Collection
    
    For Each prop In dbInfo_(entity_name)(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" Then Call col.Add(prop(DBINFO_PROPERTY_NAME))
    Next
    
    Set GetPropertyNamesFromDbInfoBy = col
End Function

Private Function GetRelatedPropFromDbInfoBy(entity_name As String) As Collection
    Dim col As New Collection, prop As Collection
    
    For Each prop In dbInfo_(entity_name)(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) = "" Then Call col.Add(prop)
    Next
    
    Set GetRelatedPropFromDbInfoBy = col
End Function

Private Function GetDbInfoSqlFieldNameBy(entity_name As String, property_name As String) As String
    GetDbInfoSqlFieldNameBy = dbInfo_(entity_name)(DBINFO_SQL_TABLE_NAME) & "." & dbInfo_(entity_name)(DBINFO_PROPERTIES)(property_name)(DBINFO_FIELD_NAME)
End Function







' Conditions -------------------------------------------------------------------------
Public Function CanGetAdoxKey() As Boolean
   CanGetAdoxKey = InStr(cn_.PROVIDER, "SQL") = 0
End Function

Private Function IsRelatedPropertyUniqueKey(prop) As Boolean
    Dim rp, re
    
    re = prop(DBINFO_PROPERTY_TYPE_NAME)
    rp = prop(DBINFO_RELATED_PROPERTY_NAME)
    
    IsRelatedPropertyUniqueKey = dbInfo_(re)(DBINFO_PROPERTIES)(rp)(DBINFO_IS_UNIQUE_KEY)
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

Private Function ConvertRecordRowsToCollection(rsw) As Collection
    Dim col As New Collection, i As Long, j As Long
    For j = 1 To UBound(rsw, 2) + 1
        Dim arr: ReDim arr(1 To UBound(rsw, 1) + 1)
        For i = 1 To UBound(rsw, 1) + 1
            arr(i) = rsw(i - 1, j - 1)
        Next i
        Call col.Add(arr, CStr(j))
    Next j
    Set ConvertRecordRowsToCollection = col
End Function

Private Function ConvertCollectionToDictionaryWithIndexKey(col As Collection) As Dictionary
    Dim dic As New Dictionary, c, i As Long
    i = 1
    For Each c In col
        Call dic.Add(CStr(i), c)
        i = i + 1
    Next
    Set ConvertCollectionToDictionaryWithIndexKey = dic
End Function

Private Function IsNullOrEmptyOrZeroOrWhiteSpace(v As Variant) As Boolean
    IsNullOrEmptyOrZeroOrWhiteSpace = True
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If vbInteger <= VarType(v) And VarType(v) <= vbCurrency Then If (v) = 0 Then Exit Function
    If VarType(v) = vbString Then If (v) = "" Then Exit Function
    IsNullOrEmptyOrZeroOrWhiteSpace = False
End Function


