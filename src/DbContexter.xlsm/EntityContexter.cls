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
' Version   : 0.5.3
' Project   : https://github.com/yyukki5/EntityContexter
' Depend on : Microsoft ActiveX Data Objects X.X Library, Microsoft ADO Ext. 2.8 for DDL and Security, Microsoft Scripting RunTime
' Before Using : Need to check "Trust access to the VBA project object model"
' Function     : Scaffolding DB, Update, Insert, Delete
' ---------------------------------------------------------------------------------

Private Const VBA_CLASS_MODULE_ID As Long = 2

Private Const ENTIINFO_ENTITIES_NAME As String = "EntitiesName"
Private Const ENTIINFO_ENTITY_NAME As String = "EntityName"
Private Const ENTIINFO_PROPERTIES As String = "Properties"
Private Const ENTIINFO_PROPERTY_NAME As String = "PropertyName"
Private Const ENTIINFO_PROPERTY_TYPE As String = "PropertyType"
Private Const ENTIINFO_PROPERTY_TYPE_PARAMETER As String = "PropertyTypeParameter"
Private Const ENTIINFO_PROPERTY_FIELD_NAME As String = "PropertyFieldName"
Private Const ENTIINFO_BASE_PROPERTY_NAME As String = "DbBaseProperty"
Private Const ENTIINFO_RELATED_PROPERTY_NAME As String = "DbRelatedProperty"
Private Const ENTIINFO_DB_TABLE_NAME As String = "DbTableName"
Private Const ENTIINFO_DB_SQL_TABLE_NAME As String = "SQLTableName"
Private Const ENTIINFO_DB_FIELD_NAME As String = "DbFieldName"
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
Private Const DBINFO_RELATED_ENTITY_NAME As String = "PropertyTypeName"
Private Const DBINFO_FIELD_NAME As String = "DbFieldName"
Private Const DBINFO_FIELD_TYPE As String = "DbFieldType"
Private Const DBINFO_IS_AUTONUMBER As String = "DbIsAutoNumber"
Private Const DBINFO_IS_PRIMARY_KEY As String = "DbIsPrimaryKey"
Private Const DBINFO_IS_UNIQUE_KEY As String = "DbIsUniqueKey"
Private Const DBINFO_BASE_PROPERTY_NAME As String = "DbBaseProperty"
Private Const DBINFO_RELATED_PROPERTY_NAME As String = "DbRelatedProperty"

Private Const KEY_DB_RELATED_TABLE_NAME As String = "RelationBaseTableName"
Private Const KEY_DB_RELATED_FIELD_NAME As String = "RelationBaseFieldName"
Private Const KEY_DB_BASE_KEY_TABLE_NAME As String = "RelationRelatedTableName"
Private Const KEY_DB_BASE_KEY_FIELD_NAME As String = "RelationRelatedFieldName"

Private Const SQL_TABLE_AS As String = "T"


Private cn_ As ADODB.Connection
Private catalog_ As ADOX.Catalog
Private isTransactionActive_ As Boolean
Private schemaName_ As String
Private isExportFiles_ As Boolean

Private baseEntityName_ As String
Private querySelect_ As String
Private queryField_ As String
Private queryFrom_ As String
Private queryJoin_ As String
Private queryWhere_ As String
Private queryOrder_ As String
Private lastIncludedEntityName_ As String
Private includedConnectTableString_ As String
Private queryFields_ As Dictionary
Private queryFieldsByProperty_ As Dictionary
Private lastTableAsIndex_ As Long

Private dbInfo_ As Collection


Private Sub Class_Initialize()
    Set cn_ = New ADODB.Connection
    Set catalog_ = New ADOX.Catalog
    isExportFiles_ = True
End Sub

Public Function Initialize(connection_string As String) As EntityContexter
    
    cn_.ConnectionString = connection_string
    
    Set Initialize = Me
End Function

Private Sub Class_Terminate()
    DisconnectDb
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else Set cn_ = Nothing
End Sub

' Properties -------------------------------------------------------------------------------
Public Property Set DbInfo(db_info As Object)
    Set dbInfo_ = db_info
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
    
    Call SetEntityRelationship(entityInfos)
                
    ' create class module from table names
    For Each table_name In tableNames
        Call CreateEntityClassBy(entityInfos(table_name(ENTIINFO_DB_TABLE_NAME)))
    Next
    Call CreateDbContext(entityInfos, dbcontext_name)
    Debug.Print "- Done Scaffolding-------"
    
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
    Optional unique_key As Boolean = False, _
    Optional related_entity_name As String, _
    Optional related_property_name As Variant, _
    Optional base_property_name As Variant _
    )
    Dim prop As Collection
    
    Set prop = New Collection
    Call prop.Add(property_name, DBINFO_PROPERTY_NAME)
    Call prop.Add(db_field_name, DBINFO_FIELD_NAME)
    Call prop.Add(db_field_type, DBINFO_FIELD_TYPE)
    Call prop.Add(auto_increment, DBINFO_IS_AUTONUMBER)
    Call prop.Add(primary_key, DBINFO_IS_PRIMARY_KEY)
    Call prop.Add(unique_key, DBINFO_IS_UNIQUE_KEY)
    
    ' todo apply multi relationship by using array
    Call prop.Add(related_entity_name, DBINFO_RELATED_ENTITY_NAME)
    Call prop.Add(related_property_name, DBINFO_RELATED_PROPERTY_NAME)
    Call prop.Add(base_property_name, DBINFO_BASE_PROPERTY_NAME)
    
    Set CreateDbInfoPropertyBy = prop
End Function

Public Function SetQueryDbSet(entity_name) As String
    baseEntityName_ = entity_name

    includedConnectTableString_ = ""
    includedConnectTableString_ = baseEntityName_
    
    Set queryFields_ = New Dictionary
    Set queryFieldsByProperty_ = New Dictionary
    querySelect_ = "select "
    queryField_ = ""
    queryJoin_ = ""
    queryWhere_ = ""
    queryOrder_ = ""
        
    Call SetQueryField("[" & SQL_TABLE_AS & "]", dbInfo_(baseEntityName_)(DBINFO_PROPERTIES))
    queryFrom_ = "from " & dbInfo_(baseEntityName_)(DBINFO_SQL_TABLE_NAME) & " as [" & SQL_TABLE_AS & "] " & vbCrLf
    
End Function

Public Function SetQueryWhere(property_name As String, sql_operator As String, compare_value As String, ByVal prop_names As Variant)
    Dim named_field_name As String
    
    If IsArray(prop_names) = False Then Call Err.Raise(9001, "EntityContexter.SetQueryWhere()", "need array")
    named_field_name = GetNamedFieldName(property_name, prop_names)
    If queryWhere_ = "" Then
        queryWhere_ = " where " & named_field_name & " " & sql_operator & " " & compare_value
    Else
        queryWhere_ = queryWhere_ & " and " & named_field_name & " " & sql_operator & " " & compare_value
    End If
    
End Function

Public Function SetQueryOrderBy(property_name As String, ByVal prop_names As Variant, Optional is_asc As Boolean) As String
    Dim named_field_name As String
    
    If IsArray(prop_names) = False Then Call Err.Raise(9001, "EntityContexter.SetQueryOrderBy()", "need array")
    named_field_name = GetNamedFieldName(property_name, prop_names)
    queryOrder_ = " order by " & named_field_name & IIf(is_asc, " asc", " desc")
    
End Function

Public Function SetQueryThenBy(property_name As String, ByVal prop_names As Variant, Optional is_asc As Boolean) As String
    Dim named_field_name As String
    
    If IsArray(prop_names) = False Then Call Err.Raise(9001, "EntityContexter.SetQueryThenBy()", "need array")
    named_field_name = GetNamedFieldName(property_name, prop_names)
    queryOrder_ = queryOrder_ & " , " & named_field_name & IIf(is_asc, " asc", " desc")
    
End Function

Public Function SetQueryInclude(property_name As String) As String
    Dim base_field As String, related_table As String, related_field As String, base_prop_name As String, related_prop_name As String, related_entity_name As String
    Dim prop As Collection, full_name As String, qry As String, prop_type_name As String

    On Error GoTo errNotFoundProperty
    base_prop_name = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_BASE_PROPERTY_NAME)
    related_prop_name = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_RELATED_PROPERTY_NAME)
    related_entity_name = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_RELATED_ENTITY_NAME)
    base_field = dbInfo_(baseEntityName_)(DBINFO_PROPERTIES)(base_prop_name)(DBINFO_FIELD_NAME)
        
    related_table = dbInfo_(related_entity_name)(DBINFO_SQL_TABLE_NAME)
    related_field = dbInfo_(related_entity_name)(DBINFO_PROPERTIES)(related_prop_name)(DBINFO_FIELD_NAME)
        
    Dim r_table_as As String
    r_table_as = SQL_TABLE_AS & lastTableAsIndex_ + 1
    qry = " left outer join " & related_table & " as [" & r_table_as & "] on " _
        & "[" & SQL_TABLE_AS & "].[" & base_field & "] = [" & r_table_as & "].[" & related_field & "]" & vbCrLf
    
    queryJoin_ = queryJoin_ & qry
    
    includedConnectTableString_ = baseEntityName_
    includedConnectTableString_ = includedConnectTableString_ & property_name
    
    Call SetQueryField("[" & r_table_as & "]", dbInfo_(related_entity_name)(DBINFO_PROPERTIES))
    
    lastIncludedEntityName_ = related_entity_name
    lastTableAsIndex_ = lastTableAsIndex_ + 1
    Exit Function
    
errNotFoundProperty:
    Call Err.Raise(100001, Err.Source & vbCrLf & "EntityContexter.SetQueryInclude(" & property_name & ")", "Not found property in DbInfo")
End Function

Public Function SetQueryThenInclude(property_name As String) As String
    Dim base_field As String, related_table As String, related_field As String, base_prop_name As String, related_prop_name As String, related_entity_name As String
    Dim prop As Collection, full_name As String, qry As String, prop_type_name As String
    
    On Error GoTo errNotFoundProperty
    base_prop_name = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_BASE_PROPERTY_NAME)
    related_prop_name = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_RELATED_PROPERTY_NAME)
    related_entity_name = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(property_name)(DBINFO_RELATED_ENTITY_NAME)
    base_field = dbInfo_(lastIncludedEntityName_)(DBINFO_PROPERTIES)(base_prop_name)(DBINFO_FIELD_NAME)
    related_table = dbInfo_(related_entity_name)(DBINFO_SQL_TABLE_NAME)
    related_field = dbInfo_(related_entity_name)(DBINFO_PROPERTIES)(related_prop_name)(DBINFO_FIELD_NAME)
    
    Dim r_table_as As String, b_table_as As String
    b_table_as = SQL_TABLE_AS & lastTableAsIndex_
    r_table_as = SQL_TABLE_AS & lastTableAsIndex_ + 1
    qry = " left outer join " & related_table & " as [" & r_table_as & "] on " _
        & "[" & b_table_as & "].[" & base_field & "] = [" & r_table_as & "].[" & related_field & "]" & vbCrLf
            
    queryJoin_ = queryJoin_ & qry
    
    includedConnectTableString_ = includedConnectTableString_ & property_name

    Call SetQueryField("[" & r_table_as & "]", dbInfo_(related_entity_name)(DBINFO_PROPERTIES))
      
    lastIncludedEntityName_ = related_entity_name
    lastTableAsIndex_ = lastTableAsIndex_ + 1
    Exit Function
    
errNotFoundProperty:
    Call Err.Raise(100001, Err.Source & vbCrLf & "EntityContexter.SetQueryThenInclude(" & property_name & ")", "Not found property in DbInfo of " & lastIncludedEntityName_)
End Function

Public Function FetchRecords(dbc As Object)
    Dim rs As Recordset, qry As String, re As Collection, records_rows As Collection
    
    qry = querySelect_ & vbCrLf & queryField_ & vbCrLf & queryFrom_ & vbCrLf & queryJoin_ & vbCrLf & queryWhere_ & vbCrLf & queryOrder_
    
    Call OpenDb
    Set rs = FetchRecordSetByQuery(qry)
    If rs.BOF And rs.EOF Then GoTo errFetchQuery
        
    Set records_rows = ConvertRecordRowsToCollection(rs.GetRows)
    Set re = FetchRecsSubRecursive(records_rows, baseEntityName_, baseEntityName_, dbc)
    
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
    Call Err.Raise(10000, Err.Source & vbCrLf & "EntityContexter.SaveChanges()", "Failed to save changes")

End Function

Public Function Add(rec As Object)
    Dim rs As New ADODB.Recordset, prop As Collection
    Dim ent_name, is_related_field_index As Boolean, r As Object, v As Variant
                                            
On Error GoTo errOpen
    OpenDb
    cn_.CursorLocation = adUseClient
    
    Call AddRelatedEntity(rec)
                
    If GetDbInfoHasPrimaryKeyBy(TypeName(rec)) = False Then If ExistUniqueRecord(rec) Then Call Err.Raise(10003, "EntityContexter.Add()", "Record is already exist")
    Call BeginTransaction
    Call rs.Open(Source:=dbInfo_(TypeName(rec))(DBINFO_SQL_TABLE_NAME), ActiveConnection:=cn_, LockType:=adLockOptimistic)
    
    Call rs.AddNew(GetFieldsListByEntity(rec), GetValuesListByEntity(rec))
    Call SetAutoincrementFieldsValues(rs, rec, dbInfo_)
    
    Call AddRelatedEntities(rec)
    Exit Function

errOpen:
    DisconnectDb
    Call Err.Raise(Err.Number, Err.Source & vbCrLf & "EntityContexter.Add()", Err.Description)
End Function

Public Function Update(rec As Object)
    
On Error GoTo errOpen
    Call OpenDb
    cn_.CursorLocation = adUseClient
    
    Call UpdateRelatedEntity(rec)

    If GetDbInfoHasPrimaryKeyBy(TypeName(rec)) = False Then Call Err.Raise(10001, "EntityContexter.Update()", "Table has no primary key")
    If ExistUniqueRecord(rec) = False Then
        Call Add(rec)
    Else
        Call BeginTransaction
        Call cn_.Execute(CommandText:=CreateUpdateQuery(rec))
    End If
    
    Call UpdateRelatedEntities(rec)
    Exit Function

errOpen:
    DisconnectDb
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Update()", Err.Description
End Function


Public Function Remove(rec As Object)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    cn_.CursorLocation = adUseClient
    If GetDbInfoHasPrimaryKeyBy(TypeName(rec)) = False Then Call Err.Raise(10001, "EntityContexter.Remove()", "Table has no primary key")
    If ExistUniqueRecord(rec) = False Then Call Err.Raise(10002, "EntityContexter.Remove()", "Not found records")
    
    Call BeginTransaction
    Call cn_.Execute(CommandText:=CreateDeleteQuery(rec))
    Exit Function

errOpen:
    DisconnectDb
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Remove()", Err.Description
End Function




' Private Methods -------------------------------------------------------------------------------------------------

' Scadffolding() -----------------------------------------------------------------
Private Property Get ProviderHasSchema() As Boolean
    ProviderHasSchema = InStr(cn_.PROVIDER, "SQLOLEDB") > 0
End Property

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

    Call entity_info.Add(table_name_info(ENTIINFO_DB_TABLE_NAME), ENTIINFO_DB_TABLE_NAME)
    Call entity_info.Add(table_name_info(ENTIINFO_DB_SQL_TABLE_NAME), ENTIINFO_DB_SQL_TABLE_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntityClassName(table_name), ENTIINFO_ENTITY_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntitesClassName(table_name), ENTIINFO_ENTITIES_NAME)
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
    Call prop.Add(col.Name, ENTIINFO_DB_FIELD_NAME)
    
    Set GetPropertyInfoByColumn = prop
End Function

Private Function SetEntityRelationship(entity_infos As Collection) As Collection
    Dim rels As Collection, rel As Variant, prop As Collection, base_ent As Collection, base_props As Collection, base_prop As Collection, rel_table_name As String, rel_ent As Collection, rel_prop As Collection
    
    For Each rel In GetRelationsOfDataBase()
        Set prop = New Collection
        
        Set base_props = entity_infos(rel(1))(ENTIINFO_PROPERTIES)
        Set base_prop = SelectEntiInfoPropertyByTableNameColumnName(entity_infos, (rel(1)), (rel(3)))
        Set rel_ent = entity_infos(rel(2))
        Set rel_prop = SelectEntiInfoPropertyByTableNameColumnName(entity_infos, (rel(2)), (rel(4)))
        rel_table_name = rel(2)
                
        ' case related column is Unique Key     then related is entity
        ' case related column is NOT Unique Key then related is entities
        If rel_prop(ENTIINFO_DB_IS_UNIQUE_KEY) Then
            Set prop = GetPropertyInfoOfRelatedEntity(rel_table_name, base_prop(ENTIINFO_PROPERTY_NAME), rel_prop(ENTIINFO_PROPERTY_NAME))
        Else
            Set prop = GetPropertyInfoOfRelatedEntities(rel_table_name, base_prop(ENTIINFO_PROPERTY_NAME), rel_prop(ENTIINFO_PROPERTY_NAME))
        End If
                
        Call base_props.Add(prop, Key:=prop(ENTIINFO_PROPERTY_NAME))
    Next
    
End Function

Private Function GetRelationsOfDataBase() As Collection
    Dim key_infos As Collection
            
    If CanGetAdoxKey Then
        Set key_infos = GetKeyInfosOfAdox()
    Else
        Set key_infos = GetKeyInfosOfTSql()
    End If
    
    Set GetRelationsOfDataBase = key_infos
End Function

Private Function SelectEntiInfoPropertyByTableNameColumnName(ents As Collection, table_name As String, column_name As String) As Collection
    Dim prop As Collection
    For Each prop In ents(table_name)(ENTIINFO_PROPERTIES)
        If prop(ENTIINFO_DB_FIELD_NAME) = column_name Then Set SelectEntiInfoPropertyByTableNameColumnName = prop
    Next
End Function

Private Function GetKeyInfosOfAdox() As Collection
    Dim tbl As ADOX.Table, ky As ADOX.Key, col As ADOX.Column, key_infos As New Collection, key_info As New Collection
    
    Call OpenDb
    For Each tbl In catalog_.Tables
        For Each ky In tbl.Keys
            If ky.Type = adKeyForeign Then
                For Each col In ky.Columns
                    Set key_info = New Collection
                    Call key_info.Add(base_table_name, KEY_DB_BASE_KEY_TABLE_NAME)
                    Call key_info.Add(ky.RelatedTable, KEY_DB_RELATED_TABLE_NAME)
                    Call key_info.Add(col.Name, KEY_DB_BASE_KEY_FIELD_NAME)
                    Call key_info.Add(col.RelatedColumn, KEY_DB_RELATED_FIELD_NAME)
                    Call key_infos.Add(key_info, ky.Name)
                Next
            End If
        Next
    Next
    
    Set GetKeyInfosOfAdox = key_infos
End Function

Private Function GetKeyInfosOfTSql() As Collection
    Dim key_infos As New Collection, key_info As New Collection
    
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
        Call key_info.Add(rs(1, i), KEY_DB_BASE_KEY_TABLE_NAME)
        Call key_info.Add(rs(3, i), KEY_DB_RELATED_TABLE_NAME)
        Call key_info.Add(rs(2, i), KEY_DB_BASE_KEY_FIELD_NAME)
        Call key_info.Add(rs(4, i), KEY_DB_RELATED_FIELD_NAME)
        Call key_infos.Add(key_info)
    
        Set key_info = New Collection
        Call key_info.Add(rs(3, i), KEY_DB_BASE_KEY_TABLE_NAME)
        Call key_info.Add(rs(1, i), KEY_DB_RELATED_TABLE_NAME)
        Call key_info.Add(rs(4, i), KEY_DB_BASE_KEY_FIELD_NAME)
        Call key_info.Add(rs(2, i), KEY_DB_RELATED_FIELD_NAME)
        Call key_infos.Add(key_info)
    Next

    Set GetKeyInfosOfTSql = key_infos
End Function

Private Function GetPropertyInfoOfRelatedEntity(related_table_name As String, base_property_name As String, related_property_name As String)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String, propertyTypeParameterName As String

    propertyName = ConvertDbTableNameToEntityClassName(related_table_name)
    propertyFieldName = ConvertDbTableNameToEntityClassFieldName(related_table_name)
    propertyTypeName = propertyName
    propertyTypeParameterName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, ENTIINFO_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, ENTIINFO_PROPERTY_FIELD_NAME)
    Call prop.Add(propertyTypeName, ENTIINFO_PROPERTY_TYPE)
    Call prop.Add(propertyTypeParameterName, ENTIINFO_PROPERTY_TYPE_PARAMETER)
    Call prop.Add("", ENTIINFO_DB_FIELD_NAME)
        
    Call prop.Add(base_property_name, ENTIINFO_BASE_PROPERTY_NAME)
    Call prop.Add(related_property_name, ENTIINFO_RELATED_PROPERTY_NAME)
    
    Call prop.Add(False, ENTIINFO_DB_IS_INDEX)
    Call prop.Add(False, ENTIINFO_DB_IS_PRIMARY_KEY)
    Call prop.Add(False, ENTIINFO_DB_IS_UNIQUE_KEY)
    Call prop.Add(False, ENTIINFO_DB_IS_AUTOINCREMENT)
    
    Set GetPropertyInfoOfRelatedEntity = prop
End Function

Private Function GetPropertyInfoOfRelatedEntities(related_table_name As String, base_property_name As String, related_property_name As String)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String, propertyTypeParameterName As String

    propertyName = ConvertDbTableNameToEntitesClassName(related_table_name)
    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(related_table_name)
    propertyTypeName = propertyName
    propertyTypeParameterName = ConvertDbTableNameToEntityClassName(related_table_name)
    
    Dim prop As New Collection
    Call prop.Add(propertyName, ENTIINFO_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, ENTIINFO_PROPERTY_FIELD_NAME)
    Call prop.Add(propertyTypeName, ENTIINFO_PROPERTY_TYPE)
    Call prop.Add(propertyTypeParameterName, ENTIINFO_PROPERTY_TYPE_PARAMETER)
    Call prop.Add("", ENTIINFO_DB_FIELD_NAME)
    
    Call prop.Add(base_property_name, ENTIINFO_BASE_PROPERTY_NAME)
    Call prop.Add(related_property_name, ENTIINFO_RELATED_PROPERTY_NAME)
      
    Call prop.Add(False, ENTIINFO_DB_IS_INDEX)
    Call prop.Add(False, ENTIINFO_DB_IS_PRIMARY_KEY)
    Call prop.Add(False, ENTIINFO_DB_IS_UNIQUE_KEY)
    Call prop.Add(False, ENTIINFO_DB_IS_AUTOINCREMENT)
    
    Set GetPropertyInfoOfRelatedEntities = prop
End Function

Private Function GetIsAutoNumberField(col As ADOX.Column, table_name As String) As Boolean
    Dim qry As String
    
    If ProviderHasSchema Then
        qry = "SELECT c.name AS ColumnName " & "FROM sys.columns c " & "INNER JOIN sys.tables t ON c.object_id = t.object_id " & "INNER JOIN sys.schemas s ON t.schema_id = s.schema_id " & _
             "WHERE t.name = '" & table_name & "' " & "AND s.name = '" & schemaName_ & "' " & "AND c.is_identity = 1;"
        GetIsAutoNumberField = Replace(cn_.Execute(qry).GetString(RowDelimeter:=","), ",", "") = col.Name
    Else
        GetIsAutoNumberField = col.Properties("Autoincrement")
    End If
    
End Function

Private Function CreateDbContext(entity_infos As Collection, dbcontext_name As String)
    Dim codeDbcClass As String
    
    codeDbcClass = GetDbContextCode(entity_infos, dbcontext_name)
    If isExportFiles_ Then
        Call AddClassModuleFromFile("DBContext", codeDbcClass)
    Else
    End If

End Function

Private Function CreateEntityClassBy(entityInfos As Collection)
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
    Call Kill(classFilePath)
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

Private Function GetVBComponent(component_name As String) As Object
    Dim c
    
    Set GetVBComponent = Nothing
    With Application.VBE.ActiveVBProject
        For Each c In .VBComponents
            If c.Name = component_name Then
                Set GetVBComponent = c
                Exit Function
            End If
        Next
    End With
End Function

Private Function GetDbContextCode(infos As Collection, dbc_name As String) As String
    Dim texts As New Collection, str As String, txt, tmp As String, info As Collection, prop As Collection
    
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
        .Add "    Set ec_.DbInfo = DbInfo"
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
        .Add "Public Function Where(property_name As String, sql_operator As String, compare_value As String, ParamArray prop_names()) As " & dbc_name
        .Add "        Call ec_.SetQueryWhere(property_name, sql_operator, compare_value, prop_names):      Set Where = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function Include(entity_name As String) As " & dbc_name
        .Add "    Call ec_.SetQueryInclude(entity_name):     Set Include = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenInclude(entity_name As String) As TableDbContextTest"
        .Add "    Call ec_.SetQueryThenInclude(entity_name):      Set ThenInclude = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function OrderBy(property_name As String, ParamArray prop_names()) As " & dbc_name
        .Add "    Call ec_.SetQueryOrderBy(property_name, prop_names, true):   Set OrderBy = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function OrderByDescending(property_name As String, ParamArray prop_names()) As " & dbc_name
        .Add "    Call ec_.SetQueryOrderBy(property_name, prop_names, false):   Set OrderByDescending = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenBy(property_name As String, ParamArray prop_names()) As " & dbc_name
        .Add "    Call ec_.SetQueryThenBy(property_name, prop_names, false):   Set ThenBy = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ThenByDescending(property_name As String, ParamArray prop_names()) As " & dbc_name
        .Add "    Call ec_.SetQueryThenBy(property_name, prop_names, false):   Set ThenByDescending = Me"
        .Add "End Function"
        .Add ""
        .Add "Public Function ToCollection() As Variant"
        .Add "    Set ToCollection = ec_.FetchRecords(Me)"
        .Add "End Function"
        .Add ""
        .Add ""
        .Add "Public Function Add(rec As Object)"
        .Add ""
        .Add "    Call ec_.Add(rec)"
        .Add "    Call ec_.SaveChanges()"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Update(rec As Object)"
        .Add ""
        .Add "    Call ec_.Update(rec)"
        .Add "    Call ec_.SaveChanges()"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Remove(rec As Object)"
        .Add ""
        .Add "    Call ec_.Remove(rec)"
        .Add "    Call ec_.SaveChanges()"
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
                If prop(ENTIINFO_DB_FIELD_NAME) = "" Then
                    tmp = tmp & ", related_entity_name:=""" & prop(ENTIINFO_PROPERTY_TYPE_PARAMETER) & """"
                    tmp = tmp & ", related_property_name:=""" & prop(ENTIINFO_RELATED_PROPERTY_NAME) & """"
                    tmp = tmp & ", base_property_name:=""" & prop(ENTIINFO_BASE_PROPERTY_NAME) & """"
                End If
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
            If prop(ENTIINFO_DB_FIELD_NAME) <> "" Then
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

Private Function ConvertDbTableNameToEntitesClassName(dbTbleName As String) As String
    ConvertDbTableNameToEntitesClassName = GetEstimatedPluralName(ConvertSnakeToPascalCase(ConvertToCodeChar(dbTbleName)))
End Function

Private Function ConvertDbTableNameToEntityClassName(dbTbleName As String) As String
    ConvertDbTableNameToEntityClassName = GetEstimatedSingularName(ConvertSnakeToPascalCase(ConvertToCodeChar(dbTbleName)))
End Function

Private Function ConvertDbTableNameToEntityClassFieldName(dbTbleName As String) As String
    ConvertDbTableNameToEntityClassFieldName = GetEstimatedSingularName(ConvertSnakeToCamelCase(ConvertToCodeChar(dbTbleName))) & "_"
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
                Call queryFieldsByProperty_.Add(full_name, table_name & ".[" & prop(DBINFO_FIELD_NAME) & "]")
            End If
        End If
    Next
    
    queryField_ = Left(queryField_, Len(queryField_) - Len(", "))
End Function

Private Function ConvertDbInfoFullNameToColumnName(full_field_name As String) As String
    ConvertDbInfoFullNameToColumnName = Replace(Replace(Replace(full_field_name, ".", ""), "[", ""), "]", "")
End Function

Private Function GetNamedFieldName(property_name As String, ByVal prop_names As Variant)
    Dim prop_path As String, prop_name
     
    prop_path = baseEntityName_
    For Each prop_name In prop_names
        prop_path = prop_path & prop_name
    Next
    
    GetNamedFieldName = queryFieldsByProperty_.Item(prop_path & property_name)
End Function


' FetchRecords() functions ------------------------------------------------------------------------------------------------------------------------
Private Function FetchRecordSetByQuery(query_string As String) As ADODB.Recordset
    OpenDb
    Set FetchRecordSetByQuery = New ADODB.Recordset
    FetchRecordSetByQuery.CursorLocation = adUseClient
    Call FetchRecordSetByQuery.Open(query_string, cn_, LockType:=adLockReadOnly, CursorType:=adOpenStatic)
End Function


Private Function FetchRecsSubRecursive(records_row As Collection, ByVal entity_name As String, ByVal prop_path As String, dbc) As Object
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
                Set re = FetchRecsSubRecursive(record_rows, prop(DBINFO_RELATED_ENTITY_NAME), prop_path & prop_name, dbc)
                
                If re.Count = 0 Then
                    Call CallByName(rec, prop_name, VbSet, Nothing)
                Else
                    If IsRelatedPropertyUniqueKey(prop) Then
                        Call CallByName(rec, prop_name, VbSet, re(1))
                    Else
                        Set entities_object = dbc.CreateObjectBy((dbInfo_(prop(DBINFO_RELATED_ENTITY_NAME))(DBINFO_ENTITIES_NAME)))
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



' Add() ----------------------------------------------------------------------------------------------------------
Private Function AddRelatedEntity(rec As Object)
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean, r As Object, v As Variant
    Dim related_entity_properties As New Collection
    
    Set related_entity_properties = GetRelatedEntityEntitiesProperty(rec, is_entity:=True)
    
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

Private Function AddRelatedEntities(rec As Object)
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean, r As Object, v As Variant
    Dim related_entities_properties As New Collection

    Set related_entities_properties = GetRelatedEntityEntitiesProperty(rec, is_entity:=False)
    
    For Each prop In related_entities_properties
        For Each r In CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet)
            v = CallByName(rec, prop(DBINFO_BASE_PROPERTY_NAME), VbGet)
            Call CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbLet, (v))
            Call Add(r)
        Next
    Next
    
End Function

Private Function GetRelatedEntityEntitiesProperty(rec As Object, is_entity As Boolean) As Collection
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean
    Dim related_properties As New Collection

    For Each prop In dbInfo_(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) = "" Then
            ent_name = GetEntityName(rec, prop)
            If ent_name <> "Nothing" Then
                is_related_field_index = dbInfo_(ent_name)(DBINFO_PROPERTIES)(prop(DBINFO_RELATED_PROPERTY_NAME))(DBINFO_IS_UNIQUE_KEY)
                If is_related_field_index = is_entity Then
                    Call related_properties.Add(prop)
                End If
            End If
        End If
    Next
    Set GetRelatedEntityEntitiesProperty = related_properties
End Function

Private Function GetEntityName(ent As Object, prop As Collection) As String
    Dim p, ent_name
    Set p = CallByName(ent, prop(DBINFO_PROPERTY_NAME), VbGet)
    ent_name = GetDbInfoEntityNameBy(TypeName(p))
    If ent_name <> "" Then Else ent_name = TypeName(p)
    GetEntityName = ent_name
End Function

Private Function GetValuesListByEntity(rec As Object)
    Dim prop As Collection, col As New Collection
    
    For Each prop In dbInfo_(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" And prop(DBINFO_IS_AUTONUMBER) = False Then Call col.Add(CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet))
    Next
    GetValuesListByEntity = CollectionToArray(col)

End Function

Private Function GetFieldsListByEntity(rec As Object)
    Dim prop As Collection, col As New Collection

    For Each prop In dbInfo_(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" And prop(DBINFO_IS_AUTONUMBER) = False Then Call col.Add(prop(DBINFO_FIELD_NAME))
    Next
    GetFieldsListByEntity = CollectionToArray(col)

End Function

Private Function SetAutoincrementFieldsValues(rs As ADODB.Recordset, rec As Object, DbInfo As Object)
    Dim prop As Collection

    For Each prop In DbInfo(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_IS_AUTONUMBER) = True Then _
            Call CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbLet, rs.Fields(prop(DBINFO_FIELD_NAME)).Value)
    Next
End Function


Private Function ExistUniqueRecord(rec As Object) As Boolean
    Dim qry As String
    qry = "select * from " & dbInfo_(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & GetQueryPrimaryWhere(rec)
    ExistUniqueRecord = cn_.Execute(CommandText:=qry).RecordCount = 1
End Function

Private Function GetQueryPrimaryWhere(rec As Object) As String
    Dim qry As String, prop As Collection
    
    qry = " Where "
    For Each prop In dbInfo_(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_IS_PRIMARY_KEY) Then
            qry = qry & dbInfo_(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & ".[" & prop(DBINFO_FIELD_NAME) & "] = " & ConvertValueToSqlQuery(prop(DBINFO_PROPERTY_NAME), rec) & ", "
        End If
    Next
    qry = Left(qry, Len(qry) - 2)
    
    GetQueryPrimaryWhere = qry
End Function

'Private Function CreateInsertQuery(rec As Object) As String
'End Function


'Update functions -------------------------------------------------------------------------------------------------
Private Function UpdateRelatedEntity(rec As Object)
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean, r As Object, v As Variant
    Dim related_entity_properties As New Collection
    
    Set related_entity_properties = GetRelatedEntityEntitiesProperty(rec, is_entity:=True)
    
    For Each prop In related_entity_properties
        Set r = CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet)
        Dim pky_value: pky_value = CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbGet)
        If IsNullOrEmptyOrZeroOrWhiteSpace(pky_value) Then
            Call Add(r)
            v = CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbGet)
            Call CallByName(rec, prop(DBINFO_BASE_PROPERTY_NAME), VbLet, (v))
        Else
            Call Update(r)
        End If
    Next
End Function

Private Function UpdateRelatedEntities(rec As Object)
    Dim prop As Collection, ent_name As String, is_related_field_index As Boolean, r As Object, v As Variant
    Dim related_entities_properties As New Collection

    Set related_entities_properties = GetRelatedEntityEntitiesProperty(rec, is_entity:=False)
    
    For Each prop In related_entities_properties
        For Each r In CallByName(rec, prop(DBINFO_PROPERTY_NAME), VbGet)
            Dim pky_prop_name:
            pky_prop_name = GetDbInfoPrimaryKeyPropertyBy(TypeName(r))(DBINFO_PROPERTY_NAME)
            Dim pky_value: pky_value = CallByName(r, pky_prop_name, VbGet)
            
            If IsNullOrEmptyOrZeroOrWhiteSpace(pky_value) Then
                v = CallByName(rec, prop(DBINFO_BASE_PROPERTY_NAME), VbGet)
                Call CallByName(r, prop(DBINFO_RELATED_PROPERTY_NAME), VbLet, (v))
                Call Add(r)
            Else
                Call Update(r)
            End If
        Next
    Next
    
End Function

Private Function CreateUpdateQuery(rec As Object) As String
    Dim table_name As String, key_prop_name As String, key_field_name As String, key_value As Variant, qry As String, prop As Collection
    
    qry = "Update " & dbInfo_(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & " set "
    For Each prop In dbInfo_(TypeName(rec))(DBINFO_PROPERTIES)
        If prop(DBINFO_FIELD_NAME) <> "" Then
            If prop(DBINFO_IS_AUTONUMBER) = False Then _
                qry = qry & "[" & prop(DBINFO_FIELD_NAME) & "] = " & ConvertValueToSqlQuery((prop(DBINFO_PROPERTY_NAME)), rec) & ", "
        End If
    Next
    qry = Left(qry, Len(qry) - 2) & GetQueryPrimaryWhere(rec)
    
    CreateUpdateQuery = qry
End Function


' Remove() ------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateDeleteQuery(rec As Object) As String
    Dim table_name As String, key_prop_name As String, key_field_name As String, key_value As Variant
    
    ' todo update to delete related table's record by using inner join.
    CreateDeleteQuery = "Delete From " & dbInfo_(TypeName(rec))(DBINFO_SQL_TABLE_NAME) & GetQueryPrimaryWhere(rec)
    
End Function



' connections -------------------------------------------------------------------------------------------------------------------------------------
Private Function OpenDb()
    If cn_.State <> adStateOpen Then cn_.Open: Set catalog_.ActiveConnection = cn_
End Function

Private Function DisconnectDb()
    Call RollbackTransaction
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Function

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


' Get value from ADOX informaiton --------------------------------------------------------------------------------------------------------------------

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

' Get value from EntityInfo informaiton --------------------------------------------------------------------------------------------------------------------


' Get value from DBInfo informaiton --------------------------------------------------------------------------------------------------------------------
Private Function GetDbInfoEntityNameBy(Optional entities_name As String)
    Dim db_info, prop
    
    If entities_name <> "" Then
        For Each db_info In dbInfo_
            If db_info(DBINFO_ENTITIES_NAME) = entities_name Then GetDbInfoEntityNameBy = db_info(DBINFO_ENTITY_NAME): Exit Function
        Next
    End If
    
End Function

Private Function GetDbInfoEntityNameByTypeName(type_name As String) As String
    Dim ent_name As String
    
    ent_name = GetDbInfoEntityNameBy(type_name:=type_name)
    If ent_name = "" Then GetDbInfoEntityNameByTypeName = type_name Else GetDbInfoEntityNameByTypeName = ent_name
        
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

Private Function GetDbInfoPrimaryKeyPropertyBy(entity_name As String) As Collection
    Dim prop As Collection
    
    For Each prop In dbInfo_(entity_name)(DBINFO_PROPERTIES)
        If prop(DBINFO_IS_PRIMARY_KEY) Then Set GetDbInfoPrimaryKeyPropertyBy = prop: Exit Function
    Next

End Function

Private Function GetDbInfoHasPrimaryKeyBy(entity_name As String) As Boolean
    Dim prop As Collection
    
    For Each prop In dbInfo_(entity_name)(DBINFO_PROPERTIES)
        If prop(DBINFO_IS_PRIMARY_KEY) Then GetDbInfoHasPrimaryKeyBy = True
    Next
    
End Function

Private Function ConvertValueToSqlQuery(prop_name As String, rec As Object) As String
    Dim v As Variant, field_type As ADOX.DataTypeEnum
    v = CallByName(rec, prop_name, VbGet)
    field_type = dbInfo_(TypeName(rec))(DBINFO_PROPERTIES)(prop_name)(DBINFO_FIELD_TYPE)
    Select Case VarType(v)
        Case vbEmpty:  v = "Null"
        Case vbNull:   v = "Null"
        Case vbBoolean: v = IIf(v, 1, 0)
        Case vbString: v = "N'" & v & "'"
        Case vbDate
            Select Case field_type
                Case ADOX.DataTypeEnum.adDBDate: v = Format(v, "'yyyy-MM-dd'")
                Case ADOX.DataTypeEnum.adDBTime: v = Format(v, "'hh:mm:ss'")
                Case Else: v = Format(v, "'yyyy-MM-dd hh:mm:ss'")
            End Select
    End Select
    
    ConvertValueToSqlQuery = v
End Function


' Conditions -------------------------------------------------------------------------
Public Function CanGetAdoxKey() As Boolean
   CanGetAdoxKey = InStr(cn_.PROVIDER, "SQL") = 0
End Function

Private Function IsRelatedPropertyUniqueKey(prop) As Boolean
    Dim rp, re
    
    re = prop(DBINFO_RELATED_ENTITY_NAME)
    rp = prop(DBINFO_RELATED_PROPERTY_NAME)
    
    IsRelatedPropertyUniqueKey = dbInfo_(re)(DBINFO_PROPERTIES)(rp)(DBINFO_IS_UNIQUE_KEY)
End Function


' Useful functions -------------------------------------------------------------------

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


