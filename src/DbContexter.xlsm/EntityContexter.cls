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
' Version   : 0.3.0
' Project   : https://github.com/yyukki5/EntityContexter
' Depend on : Microsoft ActiveX Data Objects X.X Library, Microsoft ADO Ext. 2.8 for DDL and Security, Microsoft Scripting RunTime
' Before Using : Need to check "Trust access to the VBA project object model"
' Function     : Scaffolding DB, Update, Insert, Delete
' ---------------------------------------------------------------------------------

Private Const VBA_CLASS_MODULE_ID As Long = 2
Private Const LOOP_MAX As Long = 2

Private Const EC_KEY_ENTITY_NAME As String = "EntityName"
Private Const EC_KEY_ENTITIES_NAME As String = "EntitiesName"
Private Const EC_KEY_ENTITY_ADOX_TABLE As String = "ADOX.Table"
Private Const EC_KEY_DB_TABLE_NAME As String = "DBTableName"
Private Const EC_KEY_SQL_TABLE_NAME As String = "SqlTableName"
Private Const EC_KEY_TABLE_PRIMARY_KEY As String = "PrimaryKey"

Private Const EC_KEY_ENTITY_PROPERTIES As String = "Properties"
Private Const EC_KEY_ENTITY_PROPERTY_NAME As String = "PropertyName"
Private Const EC_KEY_ENTITY_FIELD_NAME As String = "EntityFieldName"
Private Const EC_KEY_ENTITY_PROPERTY_TYPE As String = "EntityPropertyType"
Private Const EC_KEY_ENTITY_IS_PROPERTY_RELATED As String = "RelatedProperty"
Private Const EC_KEY_ENTITY_ADOX_COLUMN As String = "ADOX.Column"
Private Const EC_KEY_ENTITY_EXIST_FOREIGNKEY As String = "ExistForeignKey"
Private Const EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY As String = "ExistInverseForeignKey"
Private Const EC_KEY_ENTITY_FOREIGNKEY As String = "ForeignKey"
Private Const EC_KEY_ENTITY_FOREIGN_TABLE As String = "ForeignTable"
Private Const EC_KEY_ENTITY_INVERSE_FOREIGNKEY As String = "InverseForeignKey"
Private Const EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE As String = "InverseForeignTable"
Private Const EC_KEY_DB_FIELD_NAME As String = "DbFieldName"
Private Const EC_KEY_DB_AUTOINCREMENT As String = "Autoincrement"

Private cn_ As ADODB.Connection
Private catalog_ As ADOX.Catalog
Private isExportFiles_ As Boolean
Private schemaName_ As String
Private connectionString_ As String

Private fetchQuery_ As String
Private fetchEntityName_ As String
Private lastSubQueryName_ As String


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
    If cn_.State <> adStateOpen Then cn_.Open
    Set catalog_.ActiveConnection = cn_
    isExportFiles_ = usingAttribute
    schemaName_ = schema_name
    
    ' get table names
    Set tableNames = GetTableNameFromDb(find_word)
    
    ' create class module code from ADOX.catalog by table name
    For Each table_name In tableNames
        Call entityInfos.Add(GetEntityInfoByTableName(table_name), table_name(EC_KEY_DB_TABLE_NAME))
    Next
    
    ' Add relationship to class module code
    Call SetEntityInfoRelatedColumns(entityInfos)
                
    ' create class module from table names
    For Each table_name In tableNames
        Call CreateEntityClassFromTableName(entityInfos(table_name(EC_KEY_DB_TABLE_NAME)))
'        Debug.Print table_name(EC_KEY_DB_TABLE_NAME) & ": done"
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
    Optional entities_name As String, _
    Optional primary_key As String)
    Dim info As Collection
    
    Set info = New Collection
    Call info.Add(table_name, EC_KEY_DB_TABLE_NAME)
    Call info.Add(sql_table_name, EC_KEY_SQL_TABLE_NAME)
    Call info.Add(entity_name, EC_KEY_ENTITY_NAME)
    Call info.Add(entities_name, EC_KEY_ENTITIES_NAME)
    Call info.Add(primary_key, EC_KEY_TABLE_PRIMARY_KEY)
    
    Set CreateDbInfoBy = info
End Function

Public Function CreateDbInfoPropertyBy( _
    Optional property_name As String, _
    Optional entity_field_name As String, _
    Optional entity_property_type As String, _
    Optional db_table_name As String, _
    Optional db_field_name As String, _
    Optional auto_increment As Boolean = False, _
    Optional foreign_key As String, _
    Optional foreign_table As String, _
    Optional inverse_foreign_key As String, _
    Optional inverse_foreign_table As String _
    )
    Dim prop As Collection
    
    Set prop = New Collection
    Call prop.Add(property_name, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(entity_field_name, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(entity_property_type, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(db_table_name, EC_KEY_DB_TABLE_NAME)
    Call prop.Add(db_field_name, EC_KEY_DB_FIELD_NAME)
    Call prop.Add(auto_increment, EC_KEY_DB_AUTOINCREMENT)
    
    If foreign_key <> "" Then
        Call prop.Add(True, EC_KEY_ENTITY_EXIST_FOREIGNKEY)
        Call prop.Add(foreign_key, EC_KEY_ENTITY_FOREIGNKEY)
        Call prop.Add(foreign_table, EC_KEY_ENTITY_FOREIGN_TABLE)
    Else
        Call prop.Add(False, EC_KEY_ENTITY_EXIST_FOREIGNKEY)
    End If
    If inverse_foreign_key <> "" Then
        Call prop.Add(True, EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY)
        Call prop.Add(inverse_foreign_key, EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
        Call prop.Add(inverse_foreign_table, EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    Else
        Call prop.Add(False, EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY)
    End If
    
    Set CreateDbInfoPropertyBy = prop
End Function


Public Function SetQueryDbSet(db_info, entity_name) As String
    fetchEntityName_ = entity_name
    fetchQuery_ = "select * from " & db_info(entity_name)(EC_KEY_SQL_TABLE_NAME)
End Function

Public Function SetQueryWhere(db_info, property_name As String, sql_operator As String, compare_value As String, Optional entity_name As String)
    entity_name = IIf(entity_name = "", fetchEntityName_, entity_name)
    lastSubQueryName_ = lastSubQueryName_ & "T"
    fetchQuery_ = "select * from (" & fetchQuery_ & ") " & lastSubQueryName_ & " Where " & lastSubQueryName_ & "." & db_info(entity_name)(EC_KEY_ENTITY_PROPERTIES)(property_name)(EC_KEY_DB_FIELD_NAME) & " " & sql_operator & " " & compare_value
End Function

Public Function SetQueryOrderBy(db_info, property_name As String, Optional entity_name As String, Optional is_asc As Boolean) As String
    entity_name = IIf(entity_name = "", fetchEntityName_, entity_name)
    lastSubQueryName_ = lastSubQueryName_ & "T"
    fetchQuery_ = "select * from (" & fetchQuery_ & ") " & lastSubQueryName_ & " order by " & lastSubQueryName_ & "." & db_info(entity_name)(EC_KEY_ENTITY_PROPERTIES)(property_name)(EC_KEY_DB_FIELD_NAME) & IIf(is_asc, " asc", " desc")
End Function
Public Function SetQueryThenBy(db_info, property_name As String, Optional entity_name As String, Optional is_asc As Boolean) As String
    entity_name = IIf(entity_name = "", fetchEntityName_, entity_name)
    fetchQuery_ = fetchQuery_ & "," & lastSubQueryName_ & "." & db_info(entity_name)(EC_KEY_ENTITY_PROPERTIES)(property_name)(EC_KEY_DB_FIELD_NAME) & IIf(is_asc, " asc", " desc")
End Function

Public Function SetQueryInclude(db_info As Variant, entity_name, entity_property_name) As String
    Dim table_name As String, related_entity_name As String, foregin_table_name As String, key_name As String, foregin_key_name As String
    
    table_name = db_info(entity_name)(EC_KEY_DB_TABLE_NAME)
    If db_info(entity_name)("Properties")(entity_property_name)(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Then
        related_entity_name = db_info(entity_name)("Properties")(entity_property_name)(EC_KEY_ENTITY_FOREIGN_TABLE)
        foregin_table_name = db_info(related_entity_name)(EC_KEY_DB_TABLE_NAME)
        key_name = db_info(entity_name)("Properties")(related_entity_name)(EC_KEY_DB_FIELD_NAME)
        foregin_key_name = db_info(entity_name)("Properties")(related_entity_name)(EC_KEY_ENTITY_FOREIGNKEY)
    End If
    If db_info(entity_name)("Properties")(entity_property_name)(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        foregin_table_name = db_info(entity_name)("Properties")(entity_property_name)(EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
        key_name = db_info(entity_name)("Properties")(entity_property_name)(EC_KEY_DB_FIELD_NAME)
        foregin_key_name = db_info(entity_name)("Properties")(entity_property_name)(EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
    End If
    
    fetchQuery_ = "(" & fetchQuery_ & ") left outer join " & foregin_table_name & " on " & table_name & "." & key_name & " = " & foregin_table_name & "." & foregin_key_name
End Function

Public Function FetchRecords(entity_name As String, DbInfo As Object, dbc As Object, cnt As Long, Optional qry As String)
    Dim rs As Recordset, Items As New Collection, T As Object, props As Collection, prop As Collection
    Dim fkey, kyfld, foreign_table_name As String, ent_name, property_name As String, qry1, tbls As Collection
        
    If cnt > LOOP_MAX Then Set FetchRecords = Nothing:     Exit Function
        
    Set props = DbInfo(entity_name)(EC_KEY_ENTITY_PROPERTIES)
    If qry = "" Then qry = fetchQuery_ ' todo
    Set rs = FetchRecordSetByQuery(qry)
        
    If rs.BOF And rs.EOF Then Set FetchRecords = Nothing:     Exit Function
    rs.MoveFirst
    Do Until rs.EOF
    
        Set T = dbc.CreateObjectBy(entity_name)
        For Each prop In props
            If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Then
                
                fkey = prop(EC_KEY_DB_FIELD_NAME)
                kyfld = rs.Fields(prop(EC_KEY_ENTITY_FOREIGNKEY)).Value
                foreign_table_name = prop(EC_KEY_ENTITY_FOREIGN_TABLE)
                ent_name = prop(EC_KEY_ENTITY_FOREIGN_TABLE)           ' todo
                
                qry1 = "SELECT * FROM " & ent_name & " WHERE " & fkey & " = " & kyfld
                Set tbls = FetchRecords((ent_name), DbInfo, dbc, cnt + 1, (qry1))
                
                property_name = prop(EC_KEY_ENTITY_PROPERTY_NAME)
                If (tbls Is Nothing) = False Then Call CallByName(CallByName(T, property_name, VbGet), "Items", VbSet, tbls)
            
            ElseIf prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
                
                fkey = prop(EC_KEY_DB_FIELD_NAME)
                kyfld = rs.Fields(prop(EC_KEY_DB_FIELD_NAME)).Value
                foreign_table_name = prop(EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
                ent_name = prop(EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)           ' todo
                
                qry1 = "SELECT * FROM " & foreign_table_name & " WHERE " & fkey & " = " & kyfld
                Set tbls = FetchRecords((ent_name), DbInfo, dbc, cnt + 1, (qry1))
                
                property_name = prop(EC_KEY_ENTITY_PROPERTY_NAME)
                If (tbls Is Nothing) = False Then Call CallByName(T, property_name, VbSet, tbls.Item(1))
            
            Else
                property_name = prop(EC_KEY_ENTITY_PROPERTY_NAME)
                Call CallByName(T, property_name, VbLet, rs.Fields(prop(EC_KEY_DB_FIELD_NAME)).Value)
            End If
        Next
        Items.Add T
        rs.MoveNext
    Loop
    rs.Close
    
    Set FetchRecords = Items
End Function

Public Function Insert(rec As Object, DbInfo As Object)
    Dim rs As New ADODB.Recordset, fl, vl
    Dim exceptional_field As String, exceptional_prop As Collection
        
    fl = GetFieldsListByEntity(rec, DbInfo)
    vl = GetValuesListByEntity(rec, DbInfo)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    Call rs.Open(Source:=GetTableNameByEntity(rec, DbInfo), ActiveConnection:=cn_, LockType:=adLockOptimistic)

    Call cn_.BeginTrans
On Error GoTo errTrans
    Call rs.AddNew(fl, vl)
    Call SetAutoAutoincrementValues(rec, DbInfo, rs)
    Call rs.Close
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
errOpen:
    If cn_.State <> adStateClosed Then Call cn_.Close
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Insert()", Err.Description
End Function

Public Function Update(rec As Object, DbInfo As Object)
    Dim rs As New ADODB.Recordset, fl, vl
    
    If HasPrimaryKey(rec, DbInfo) = False Then Call Err.Raise(10001, "EntityContexter.Update()", "Table has no primary key")
    
    fl = GetFieldsListByEntity(rec, DbInfo)
    vl = GetValuesListByEntity(rec, DbInfo)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    Call rs.Open(Source:=CreateQueryForUpdate(rec, DbInfo), ActiveConnection:=cn_, LockType:=adLockOptimistic)
        
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call rs.Update(fl, vl)
    Call rs.Close
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
errOpen:
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Update()", Err.Description
End Function

Public Function Delete(rec As Object, DbInfo As Object)
    Dim rs As New ADODB.Recordset
    
    If HasPrimaryKey(rec, DbInfo) = False Then Call Err.Raise(10001, "EntityContexter.Update()", "Table has no primary key")
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
        
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call cn_.Execute(CommandText:=CreateDeleteQuery(rec, DbInfo))
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
errOpen:
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Delete()", Err.Description
End Function

Private Sub MoveToRecord(rs As ADODB.Recordset, rec As Object, DbInfo As Collection)
    Dim index_field As String, index_value, prop As Collection

    index_field = DbInfo(TypeName(rec))(EC_KEY_TABLE_PRIMARY_KEY)
    For Each prop In DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_FIELD_NAME) = index_field Then index_value = CallByName(rec, prop(EC_KEY_ENTITY_PROPERTY_NAME), VbGet)
        End If
    Next
    
    rs.MoveFirst
    Do Until rs.EOF
        If rs.Fields(index_field).Value = index_value Then Exit Do
        rs.MoveNext
    Loop
End Sub

Private Function GetValuesListByEntity(rec As Object, DbInfo As Object)
    Dim i As Long, value_list, props As Collection, prop As Collection, col As New Collection
        
    Set props = DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)
    
    For Each prop In props
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_AUTOINCREMENT) = False Then col.Add CallByName(rec, prop(EC_KEY_ENTITY_PROPERTY_NAME), VbGet)
        End If
    Next
    GetValuesListByEntity = CollectionToArray(col)

End Function

Private Function GetFieldsListByEntity(rec As Object, DbInfo As Object)
    Dim i As Long, value_list, props As Collection, prop As Collection, col As New Collection
        
    Set props = DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)
    
    For Each prop In props
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_AUTOINCREMENT) = False Then col.Add prop(EC_KEY_DB_FIELD_NAME)
        End If
    Next
    GetFieldsListByEntity = CollectionToArray(col)
        
End Function

Private Function SetAutoAutoincrementValues(rec As Object, DbInfo As Object, rs As ADODB.Recordset)
    Dim i As Long, prop As Collection
        
    i = 0
    For Each prop In DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)
        i = i + 1
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_AUTOINCREMENT) = True Then
                Call CallByName(rec, prop(EC_KEY_ENTITY_PROPERTY_NAME), VbLet, rs.Fields(prop(EC_KEY_DB_FIELD_NAME)).Value)
            End If
        End If
    Next
End Function


Private Function HasPrimaryKey(rec As Object, DbInfo As Object) As Boolean
    
    HasPrimaryKey = DbInfo(TypeName(rec))(EC_KEY_TABLE_PRIMARY_KEY) <> ""
    
End Function

Private Function CreateDeleteQuery(rec As Object, DbInfo As Object) As String
    Dim table_name As String, key_prop_name As String, key_field_name As String, key_value As Variant
    
    ' todo update to delete related table's record by using inner join.
    table_name = GetTableNameByEntity(rec, DbInfo)
    key_prop_name = GetPrimaryKeyPropertyByEntity(rec, DbInfo)
    key_field_name = DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)(key_prop_name)(EC_KEY_DB_FIELD_NAME)
    key_value = ConvertValueToSqlCode(key_prop_name, rec, DbInfo)
    
    CreateDeleteQuery = "Delete From " & table_name & " Where " & key_field_name & " = " & key_value
    
End Function


Private Function CreateQueryForUpdate(rec As Object, DbInfo As Object) As String
    Dim table_name As String, key_prop_name As String, key_field_name As String, key_value As Variant
    
    table_name = GetTableNameByEntity(rec, DbInfo)
    key_prop_name = GetPrimaryKeyPropertyByEntity(rec, DbInfo)
    key_field_name = DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)(key_prop_name)(EC_KEY_DB_FIELD_NAME)
    key_value = ConvertValueToSqlCode(key_prop_name, rec, DbInfo)
    
    CreateQueryForUpdate = "Select * From " & table_name & " Where " & key_field_name & " = " & key_value
    
End Function

Private Function ConvertValueToSqlCode(prop_name As String, rec As Object, DbInfo As Object) As String
    ' todo, convert value to sql code using by dbtype
    Dim v As Variant
    v = CallByName(rec, prop_name, VbGet)
    If VarType(v) = vbString Then v = "'" & v & "'"
    If VarType(v) = vbDate Then v = Format(v, "yyyy-MM-DD hh:mm:ss")
    ConvertValueToSqlCode = v
End Function


Private Function GetTableNameByEntity(rec As Object, DbInfo As Object) As String
    Dim table_name As String, info As Collection
    
    For Each info In DbInfo
        If info(EC_KEY_ENTITY_NAME) = TypeName(rec) Then table_name = info(EC_KEY_SQL_TABLE_NAME)
    Next
    
    GetTableNameByEntity = table_name
End Function

Private Function GetPrimaryKeyPropertyByEntity(rec As Object, DbInfo As Object) As String
    Dim prop As Collection
    
    For Each prop In DbInfo(TypeName(rec))(EC_KEY_ENTITY_PROPERTIES)
        If prop(EC_KEY_DB_FIELD_NAME) = DbInfo(TypeName(rec))(EC_KEY_TABLE_PRIMARY_KEY) Then GetPrimaryKeyPropertyByEntity = prop(EC_KEY_ENTITY_PROPERTY_NAME)
    Next
    
End Function

Private Function GetDBInfoPropertyByFieldName(field_name As String, table_name As String, DbInfo As Object) As Collection
    Dim prop As Collection
    
    For Each prop In DbInfo(table_name)("Properties")
        If prop("DbFieldName") = field_name Then Set GetDBInfoPropertyByFieldName = prop
    Next
    
End Function

Private Property Get ProviderHasSchema() As Boolean
    ProviderHasSchema = InStr(cn_.PROVIDER, "SQLOLEDB") > 0
End Property



' Private Methods ----------------------------------------------------------------
Private Function FetchRecordSetByQuery(query_string As String) As ADODB.Recordset
    If cn_.State <> adStateOpen Then cn_.Open
    Set FetchRecordSetByQuery = New ADODB.Recordset
    FetchRecordSetByQuery.CursorLocation = adUseClient
    Call FetchRecordSetByQuery.Open(query_string, cn_, LockType:=adLockReadOnly, CursorType:=adOpenStatic)
End Function

Private Function GetTableNameFromDb(Optional find_word As String = "") As Collection
    Dim tableNames As New Collection, sch As ADODB.Recordset, table_names As Collection
    
On Error GoTo errLabel
    If cn_.State <> adStateOpen Then cn_.Open
    Set sch = cn_.OpenSchema(adSchemaTables)
    
    sch.MoveFirst
    Do Until sch.EOF
        Set table_names = New Collection
        
        If find_word <> "" And InStr(sch.Fields("TABLE_NAME").Value, find_word) = 0 Then GoTo next_loop
        
        If ProviderHasSchema Then
            If sch.Fields("TABLE_SCHEMA").Value = schemaName_ Then
                Call table_names.Add(sch.Fields("TABLE_NAME").Value, EC_KEY_DB_TABLE_NAME)
                Call table_names.Add(ConvertToSqlTableNameBy(sch.Fields("TABLE_NAME").Value, schemaName_), EC_KEY_SQL_TABLE_NAME)
                            
                Call tableNames.Add(table_names)
            End If
        Else
            If sch.Fields("TABLE_TYPE").Value = "TABLE" Or sch.Fields("TABLE_TYPE").Value = "VIEW" Then
                Call table_names.Add(sch.Fields("TABLE_NAME").Value, EC_KEY_DB_TABLE_NAME)
                Call table_names.Add(ConvertToSqlTableNameBy(sch.Fields("TABLE_NAME").Value), EC_KEY_SQL_TABLE_NAME)
                
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
    If schema_name = "" Then
        ConvertToSqlTableNameBy = "[" & table_name & "]"
    Else
        ConvertToSqlTableNameBy = "[" & schema_name & "]." & "[" & table_name & "]"
    End If
End Function
Private Function ConvertToTableNameBy(sql_table_name As String) As String
    If InStr(sql_table_name, "].[") > 0 Then
        ConvertToTableNameBy = Mid(sql_table_name, InStr(sql_table_name, "].[") + 2, Len(sql_table_name) - 1)
    Else
        ConvertToTableNameBy = Mid(sql_table_name, 2, Len(sql_table_name) - 2)
    End If
End Function




Private Function GetEntityInfoByTableName(table_name_info As Collection) As Collection
    Dim props As New Collection, prop As Collection, tbl As ADOX.Table, col As ADOX.Column, entity_info As New Collection
    Dim ky As ADOX.Key, table_name As String
    
    table_name = table_name_info(EC_KEY_DB_TABLE_NAME)
    Set tbl = catalog_.Tables(table_name)
    For Each col In tbl.Columns
        Set prop = GetPropertyInfoByColumn(col)
        Call prop.Add(GetIsAutoNumberField(col, tbl.Name), EC_KEY_DB_AUTOINCREMENT)
        Call props.Add(prop, prop(EC_KEY_ENTITY_PROPERTY_NAME))
    Next
    
    Call entity_info.Add(table_name_info(EC_KEY_DB_TABLE_NAME), EC_KEY_DB_TABLE_NAME)
    Call entity_info.Add(table_name_info(EC_KEY_SQL_TABLE_NAME), EC_KEY_SQL_TABLE_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntityClassName(table_name), EC_KEY_ENTITY_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntitesClassName(table_name), EC_KEY_ENTITIES_NAME)
    Call entity_info.Add(tbl, EC_KEY_ENTITY_ADOX_TABLE)
    Call entity_info.Add(props, EC_KEY_ENTITY_PROPERTIES)
        
    Set GetEntityInfoByTableName = entity_info
End Function


Private Function SetEntityInfoRelatedColumns(ByRef entityInfos As Collection)
    Dim EntityInfo As Collection, col As ADOX.Column, ky As ADOX.Key, tbl As ADOX.Table, prop As Collection
    
    For Each EntityInfo In entityInfos
        Set tbl = EntityInfo(EC_KEY_ENTITY_ADOX_TABLE)
        For Each ky In tbl.Keys
            If ky.Type <> adKeyForeign Then
            Else
                For Each col In ky.Columns
                    Set prop = GetPropertyInfoByInverseProperty(ky.RelatedTable, col)
                    Call EntityInfo(EC_KEY_ENTITY_PROPERTIES).Add(prop, prop(EC_KEY_ENTITY_PROPERTY_NAME))
                    Set prop = GetPropertyInfoByForeignKey(tbl.Name, col)
                    Call entityInfos(ky.RelatedTable).Item(EC_KEY_ENTITY_PROPERTIES).Add(prop, prop(EC_KEY_ENTITY_PROPERTY_NAME))
                Next
            End If
        Next
    Next
    
End Function

Private Function GetPropertyInfoByColumn(col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbFieldNameToClassPropertyName(col.Name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(col.Name)
    propertyTypeName = GetEntityPropertyTypeBy(col)

    Dim prop As New Collection
    Call prop.Add(propertyName, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(False, EC_KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col, EC_KEY_ENTITY_ADOX_COLUMN)

    Set GetPropertyInfoByColumn = prop
End Function


Private Function GetIsAutoNumberField(col As ADOX.Column, table_name As String) As Boolean
    Dim qry As String
    
    If ProviderHasSchema Then
        qry = "SELECT c.name AS ColumnName " & "FROM sys.columns c " & "INNER JOIN sys.tables t ON c.object_id = t.object_id " & "INNER JOIN sys.schemas s ON t.schema_id = s.schema_id " &
             "WHERE t.name = '" & table_name & "' " & "AND s.name = '" & schemaName_ & "' " & "AND c.is_identity = 1;"
        GetIsAutoNumberField = Replace(cn_.Execute(qry).GetString(RowDelimeter:=","), ",", "") = col.Name
    Else
        GetIsAutoNumberField = col.Properties(EC_KEY_DB_AUTOINCREMENT)
    End If
    
End Function


Private Function GetPropertyInfoByForeignKey(key_table_name As String, col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertDbTableNameToEntitesClassName(key_table_name)
    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(key_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(True, EC_KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col, EC_KEY_ENTITY_ADOX_COLUMN)
    Call prop.Add(col.Name, EC_KEY_DB_FIELD_NAME)
    Call prop.Add(col.RelatedColumn, EC_KEY_ENTITY_FOREIGNKEY)
    Call prop.Add(key_table_name, EC_KEY_ENTITY_FOREIGN_TABLE)
    Call prop.Add("", EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
    Call prop.Add("", EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    
    Set GetPropertyInfoByForeignKey = prop
End Function

Private Function GetPropertyInfoByInverseProperty(key_table_name As String, col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbTableNameToEntityClassName(key_table_name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(key_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(True, EC_KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col.RelatedColumn, EC_KEY_DB_FIELD_NAME)
    Call prop.Add("", EC_KEY_ENTITY_FOREIGNKEY)
    Call prop.Add("", EC_KEY_ENTITY_FOREIGN_TABLE)
    Call prop.Add(col.Name, EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
    Call prop.Add(key_table_name, EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    Call prop.Add(col, EC_KEY_ENTITY_ADOX_COLUMN)
    
    Set GetPropertyInfoByInverseProperty = prop
End Function


Private Function CreateDbContext(tableEntityInfos As Collection, dbcontext_name As String)
    Dim codeDbcClass As String
    
    codeDbcClass = GetDbContextCode(tableEntityInfos, dbcontext_name)
    If isExportFiles_ Then
        AddClassModuleFromFile "DBContext", codeDbcClass
    Else
    
    End If
    

End Function

Private Function CreateEntityClassFromTableName(tableEntityInfo As Collection)
    Dim EntityClass As Object, EntitiesClass As Object, codeEntityClass As String, codeEntitiesClass As String
        
    codeEntityClass = GetEntityCode(tableEntityInfo(EC_KEY_ENTITY_NAME), tableEntityInfo(EC_KEY_ENTITY_PROPERTIES))
    codeEntitiesClass = GetEntitesCode(tableEntityInfo(EC_KEY_ENTITIES_NAME), tableEntityInfo(EC_KEY_ENTITY_NAME))
    
    If isExportFiles_ Then
        ' option1 using write file and import file. to using attribute key word in class module.
        Call AddClassModuleFromFile(tableEntityInfo(EC_KEY_ENTITY_NAME), codeEntityClass)
        Call AddClassModuleFromFile(tableEntityInfo(EC_KEY_ENTITIES_NAME), codeEntitiesClass)
        
    Else
        ' option2 using CodeModule.Insertlines()
        ' only force update (delete all and re-insert)
        Set EntityClass = AddClassModule(tableEntityInfo(EC_KEY_ENTITY_NAME))
        EntityClass.CodeModule.DeleteLines 1, EntityClass.CodeModule.CountOfLines
        EntityClass.CodeModule.InsertLines 1, codeEntityClass
        
        Set EntitiesClass = AddClassModule(tableEntityInfo(EC_KEY_ENTITIES_NAME))
        EntitiesClass.CodeModule.DeleteLines 1, EntitiesClass.CodeModule.CountOfLines
        EntitiesClass.CodeModule.InsertLines 1, codeEntitiesClass
    End If
    
End Function

Private Function AddClassModuleFromFile(className, codes)
    Dim classFilePath As String
    classFilePath = ThisWorkbook.Path & "\" & className & ".cls"
    CreateClassModuleFile classFilePath, codes
    ImportClassModuleFile classFilePath, className
    DeleteClassModuleFile classFilePath
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
        .Add "Public Function Insert(rec As Object)"
        .Add ""
        .Add "    Call ec_.Insert(rec, DBInfo)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Update(rec As Object)"
        .Add ""
        .Add "    Call ec_.Update(rec, DBInfo)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Public Function Delete(rec As Object)"
        .Add ""
        .Add "    Call ec_.Delete(rec, DBInfo)"
        .Add ""
        .Add "End Function"
        .Add ""
        .Add "Private Property Get DbInfo() As Collection"
        .Add "    Dim infos As New Collection, info As Collection, props As Collection, prop As Collection"
        .Add ""
        For Each info In infos
            tmp = "    Set info = ec_.CreateDbInfoBy( "
            tmp = tmp & "  table_name    :=""" & info(EC_KEY_DB_TABLE_NAME) & """"
            tmp = tmp & ", sql_table_name:=""" & info(EC_KEY_SQL_TABLE_NAME) & """"
            tmp = tmp & ", entity_name   :=""" & info(EC_KEY_ENTITY_NAME) & """"
            tmp = tmp & ", entities_name :=""" & info(EC_KEY_ENTITIES_NAME) & """"
            tmp = tmp & ", primary_key   :=""" & GetPrimaryKeyNameByIndexes(info(EC_KEY_ENTITY_ADOX_TABLE).Indexes) & """"
            tmp = tmp & ")"
            .Add tmp
            
            .Add ""
            .Add "    Set props = New Collection"
            For Each prop In info("Properties")
                tmp = "    Set prop = ec_.CreateDbInfoPropertyBy("
                tmp = tmp & "  property_name:=""" & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """"
                tmp = tmp & ", entity_field_name:=""" & prop(EC_KEY_ENTITY_FIELD_NAME) & """"
                tmp = tmp & ", entity_property_type:=""" & prop(EC_KEY_ENTITY_PROPERTY_TYPE) & """"
                tmp = tmp & ", db_field_name:=""" & prop(EC_KEY_ENTITY_ADOX_COLUMN).Name & """"
                tmp = tmp & IIf(prop(EC_KEY_DB_AUTOINCREMENT), ", auto_increment:=" & prop(EC_KEY_DB_AUTOINCREMENT), "")
                tmp = tmp & ", foreign_key:=""" & prop(EC_KEY_ENTITY_FOREIGNKEY) & """"
                tmp = tmp & ", foreign_table:=""" & prop(EC_KEY_ENTITY_FOREIGN_TABLE) & """"
                tmp = tmp & ", inverse_foreign_key:=""" & prop(EC_KEY_ENTITY_INVERSE_FOREIGNKEY) & """"
                tmp = tmp & ", inverse_foreign_table:=""" & prop(EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE) & """"
                tmp = tmp & ")"
                .Add tmp
                .Add "    Call props.Add(prop, """ & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """)"
            Next
            .Add "    Call info.Add(props, """ & EC_KEY_ENTITY_PROPERTIES & """)"
            .Add "    Call infos.Add(info, """ & info(EC_KEY_ENTITY_NAME) & """)"
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
            .Add "        Case """ & info(EC_KEY_ENTITY_NAME) & """: Set CreateObjectBy = New " & info(EC_KEY_ENTITY_NAME)
        Next
        .Add "    End Select"
        .Add "End Function"
        .Add ""
    End With

    For Each txt In texts: str = str + txt + vbCrLf: Next
    GetDbContextCode = str
End Function

Private Function GetPrimaryKeyNameByIndexes(adox_indexes As ADOX.Indexes) As String
    Dim adox_index As ADOX.Index
    For Each adox_index In adox_indexes
        If adox_index.PrimaryKey Then GetPrimaryKeyNameByIndexes = adox_index.Columns(0).Name
    Next
End Function


Private Function GetEntitesCode(entitiesName As String, entityName As String) As String
    Dim texts As New Collection, str As String, txt

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

Private Function GetEntityCode(entityName As String, propertyInfo As Collection) As String
    Dim texts As New Collection, pInfo, str As String, txt

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
        For Each pInfo In propertyInfo
            If pInfo(EC_KEY_ENTITY_IS_PROPERTY_RELATED) = False Then
                .Add "Private " & pInfo(EC_KEY_ENTITY_FIELD_NAME) & " as " & pInfo(EC_KEY_ENTITY_PROPERTY_TYPE)
            Else
                .Add "Private " & pInfo(EC_KEY_ENTITY_FIELD_NAME) & " as New " & pInfo(EC_KEY_ENTITY_PROPERTY_TYPE)
            End If
        Next
        .Add ""
        For Each pInfo In propertyInfo
            If pInfo(EC_KEY_ENTITY_IS_PROPERTY_RELATED) = False Then
                .Add "Public Property Get " & pInfo(EC_KEY_ENTITY_PROPERTY_NAME) & "() as " & pInfo(EC_KEY_ENTITY_PROPERTY_TYPE)
                .Add "    Let " & pInfo(EC_KEY_ENTITY_PROPERTY_NAME) & " = " & pInfo(EC_KEY_ENTITY_FIELD_NAME)
                .Add "End Property"
                .Add "Public Property Let " & pInfo(EC_KEY_ENTITY_PROPERTY_NAME) & "(val as " & pInfo(EC_KEY_ENTITY_PROPERTY_TYPE) & ")"
                .Add "    Let " & pInfo(EC_KEY_ENTITY_FIELD_NAME) & " = val"
                .Add "End Property"
            Else
                .Add "Public Property Get " & pInfo(EC_KEY_ENTITY_PROPERTY_NAME) & "() as " & pInfo(EC_KEY_ENTITY_PROPERTY_TYPE)
                .Add "    Set " & pInfo(EC_KEY_ENTITY_PROPERTY_NAME) & " = " & pInfo(EC_KEY_ENTITY_FIELD_NAME)
                .Add "End Property"
                .Add "Public Property Set " & pInfo(EC_KEY_ENTITY_PROPERTY_NAME) & "(val as Variant)"
                .Add "    Set " & pInfo(EC_KEY_ENTITY_FIELD_NAME) & " = val"
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


Private Function GetEntityPropertyTypeBy(col As ADOX.Column) As String
    Dim propertyTypeName As String, fieldType As ADODB.DataTypeEnum

    fieldType = col.Type
    If col.Attributes >= adColNullable Then
        propertyTypeName = "Variant"
    ElseIf fieldType = adInteger Then propertyTypeName = "Integer"
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
    Call prop.Add("", "PrimaryKey")
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

