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
' Version   : 0.2.0
' Project   : https://github.com/yyukki5/EntityContexter
' Depend on : Microsoft ActiveX Data Objects X.X Library, Microsoft ADO Ext. 2.8 for DDL and Security
' Before Using : Need to check "Trust access to the VBA project object model"
' Function     : Scaffolding DB, Update, Insert, Delete
' ---------------------------------------------------------------------------------

Private Const VBA_CLASS_MODULE_ID As Long = 2
Private Const LOOP_MAX As Long = 5

Private Const EC_KEY_ENTITY_NAME As String = "EntityName"
Private Const EC_KEY_ENTITIES_NAME As String = "EntitiesName"
Private Const EC_KEY_ENTITY_DB_TABLE As String = "ADOX.Table"
Private Const EC_KEY_DB_TABLE_NAME As String = "DBTableName"

Private Const EC_KEY_ENTITY_PROPERTIES As String = "Properties"
Private Const EC_KEY_ENTITY_PROPERTY_NAME As String = "PropertyName"
Private Const EC_KEY_ENTITY_FIELD_NAME As String = "EntityFieldName"
Private Const EC_KEY_ENTITY_PROPERTY_TYPE As String = "EntityPropertyType"
Private Const EC_KEY_ENTITY_IS_PROPERTY_RELATED As String = "RelatedProperty"
Private Const EC_KEY_ENTITY_DB_COLUMN As String = "ADOX.Column"
Private Const EC_KEY_ENTITY_EXIST_FOREIGNKEY As String = "ExistForeignKey"
Private Const EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY As String = "ExistInverseForeignKey"
Private Const EC_KEY_ENTITY_FOREIGNKEY As String = "ForeignKey"
Private Const EC_KEY_ENTITY_FOREIGN_TABLE As String = "ForeignTable"
Private Const EC_KEY_ENTITY_INVERSE_FOREIGNKEY As String = "InverseForeignKey"
Private Const EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE As String = "InverseForeignTable"
Private Const EC_KEY_DB_FIELD_NAME As String = "DbFieldName"
Private Const EC_KEY_DB_AUTOINCREMENT As String = "Autoincrement"


' EntityInfos*
' + EntityInfo*                     Key:[EntityClassName]
' - EntityInfo*                     Key:[EntityClassName]
'   - EntityName*
'   - EntitiesName*
'   - ADOX.Table
'     - Name
'     - Indexes
'     - Keys
'     - Properties
'     - Type
'   - EntityProperties*
'     + EntityProperty*             Key:[ClassPropertyName]
'     - EntityProperty*             Key:[ClassPropertyName]
'       - EntityPropertyName*
'       - EntityFieldName*
'       - EntityPropertyType*
'       - ADOX.Column
'         - Name
'         - Properties
'         - Type
'       - RelatedProperty
'       - DbFieldName
'       - Autoincrement

Private cn_ As ADODB.Connection
Private catalog_ As ADOX.catalog
Private isExportFiles_ As Boolean
Private schemaName_ As String
Private connectionString_ As String

Public Enum PropertyInfoIndex
    [_Init]
    ecPropertyName
    ecFieldName
    ecPropertyType
    ecPropertyIsLet
    ecDbKeyFieldName
    ecDbRelatedFieldName
    [_Last]
End Enum
Public Enum EntityInfoIndex
    [_Init]
    ecDbTableName
    ecEntityClassName
    ecEntitesClassName
    ecPropertyInfo
    [_Last]
End Enum
Private Enum FieldValuesIndex
    ecField
    ecValue
End Enum


Private Sub Class_Initialize()
    Set cn_ = New ADODB.Connection
    Set catalog_ = New ADOX.catalog
    isExportFiles_ = True
End Sub

Public Function Initialize(connectionString As String) As EntityContexter
    
    connectionString_ = connectionString
    cn_.connectionString = connectionString
    catalog_.ActiveConnection = connectionString
    
    Set Initialize = Me
End Function

'Public Function SetRelationShip1toN(primaryKey As Object, primaryKeyPropertyName As String, foreignKey As Object, foreignKeyPropertyName As String)
'    Dim primaryKeyObject As Object, foreignKeyObject As Object
'
'    For Each primaryKeyObject In primaryKey.Items
'        For Each foreignKeyObject In foreignKey.Items
'            If CallByName(foreignKeyObject, foreignKeyPropertyName, VbGet) = CallByName(primaryKeyObject, primaryKeyPropertyName, VbGet) Then
'                CallByName(primaryKeyObject, TypeName(foreignKey), VbGet).Add foreignKeyObject
'                Call CallByName(foreignKeyObject, TypeName(primaryKeyObject), VbSet, primaryKeyObject)
'            End If
'        Next
'    Next
'
'End Function
'
'Public Function SetRelationShipNto1(primaryKey As Object, primaryKeyPropertyName As String, foreignKey As Object, foreignKeyPropertyName As String)
'    Dim primaryKeyObject As Object, foreignKeyObject As Object
'
'    For Each primaryKeyObject In primaryKey.Items
'        For Each foreignKeyObject In foreignKey.Items
'            If CallByName(foreignKeyObject, foreignKeyPropertyName, VbGet) = CallByName(primaryKeyObject, primaryKeyPropertyName, VbGet) Then
'                Call CallByName(primaryKeyObject, TypeName(foreignKey.Item(1)), VbSet, foreignKeyObject)
'                Call CallByName(foreignKeyObject, TypeName(primaryKey), VbSet, primaryKey)
'            End If
'        Next
'    Next
'End Function



Public Function ScaffoldDataBese(Optional usingAttribute As Boolean = True, _
                                 Optional schema_name As String, _
                                 Optional dbcontext_name As String = "DbContext" _
                                 ) As Collection
    Dim table_name, tableNames As New Collection, entityInfos As New Collection
    
    ' setting
    isExportFiles_ = usingAttribute
    schemaName_ = schema_name
    
    ' get table names
    Set tableNames = GetTableNameFromDb()
    
    ' create class module code from ADOX.catalog by table name
    For Each table_name In tableNames
        Call entityInfos.Add(GetEntityInfoByTableName((table_name)), table_name)
    Next
    
    ' Add relationship to class module code
    Call SetEntityInfoRelatedColumns(entityInfos)
                
    Debug.Print "- " & table_name & " -------"
    ' create class module from table names
    For Each table_name In tableNames
        Call CreateEntityClassFromTableName(entityInfos(table_name))
        Debug.Print table_name & ": done"
    Next
    ' to be implement:  create DbContext.cls
    Call CreateDbContext(entityInfos, dbcontext_name)
    Debug.Print "- Done -------"
    
    ' for test
    Set ScaffoldDataBese = entityInfos
End Function

Public Function CreateDbInfoBy(Optional table_name, Optional entity_name, Optional entities_name, Optional primary_key)
    Dim info As Collection
    
    Set info = New Collection
    If IsMissing(table_name) = False Then Call info.Add(table_name, EC_KEY_DB_TABLE_NAME)
    If IsMissing(entity_name) = False Then Call info.Add(entity_name, EC_KEY_ENTITY_NAME)
    If IsMissing(entities_name) = False Then Call info.Add(entities_name, EC_KEY_ENTITIES_NAME)
    If IsMissing(primary_key) = False Then Call info.Add(primary_key, "PrimaryKey")
    
    Set CreateDbInfoBy = info
End Function

Public Function CreateDbInfoPropertyBy(Optional property_name, _
    Optional entity_field_name, _
    Optional entity_property_type, _
    Optional db_table_name, _
    Optional db_field_name, _
    Optional auto_increment, _
    Optional foreign_key, _
    Optional foreign_table, _
    Optional inverse_foreign_key, _
    Optional inverse_foreign_table _
    )
    Dim prop As Collection
    
    Set prop = New Collection
    If IsMissing(property_name) = False Then Call prop.Add(property_name, EC_KEY_ENTITY_PROPERTY_NAME)
    If IsMissing(entity_field_name) = False Then Call prop.Add(entity_field_name, EC_KEY_ENTITY_FIELD_NAME)
    If IsMissing(entity_property_type) = False Then Call prop.Add(entity_property_type, EC_KEY_ENTITY_PROPERTY_TYPE)
    If IsMissing(db_table_name) = False Then Call prop.Add(db_table_name, EC_KEY_DB_TABLE_NAME)
    If IsMissing(db_field_name) = False Then Call prop.Add(db_field_name, EC_KEY_DB_FIELD_NAME)
    If IsMissing(auto_increment) = False Then Call prop.Add(auto_increment, EC_KEY_DB_AUTOINCREMENT)
    If IsMissing(foreign_key) = False Then
        Call prop.Add(True, EC_KEY_ENTITY_EXIST_FOREIGNKEY)
        Call prop.Add(foreign_key, EC_KEY_ENTITY_FOREIGNKEY)
        Call prop.Add(foreign_table, EC_KEY_ENTITY_FOREIGN_TABLE)
    Else
        Call prop.Add(False, EC_KEY_ENTITY_EXIST_FOREIGNKEY)
    End If
    If IsMissing(inverse_foreign_key) = False Then
        Call prop.Add(True, EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY)
        Call prop.Add(inverse_foreign_key, EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
        Call prop.Add(inverse_foreign_table, EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    Else
        Call prop.Add(False, EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY)
    End If
    
    Set CreateDbInfoPropertyBy = prop
End Function


Public Function FetchRecords(qry As String, entity_name As String, DbInfo As Object, dbc As Object, cnt As Long)
    Dim rs As Recordset, Items As New Collection, field As ADODB.field, propertyName As String, T As Object
    Dim table_name As String, props As Collection, prop As Collection, field_value
    Dim tbl As ADOX.Table, col As ADOX.Column
    
    Dim fkey, ent_name, qry1, tbls As Collection, kyfld, ky_name, kycol, ky As Key, T1, T2
        
    If cnt > LOOP_MAX Then Set FetchRecords = Nothing:     Exit Function
        
    table_name = DbInfo(entity_name)(EC_KEY_DB_TABLE_NAME)
    Set props = DbInfo(entity_name)(EC_KEY_ENTITY_PROPERTIES)
    Set rs = FetchRecordSetByQuery(qry)
    
    If rs.BOF And rs.EOF Then Set FetchRecords = Nothing:     Exit Function
    rs.MoveFirst
    Do Until rs.EOF
        Set T = dbc.CreateObjectBy(entity_name)
        
'        For Each field In rs.Fields
        For Each prop In props
            
            If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Then
                
                fkey = prop(EC_KEY_ENTITY_FOREIGNKEY)
                ent_name = prop(EC_KEY_ENTITY_FOREIGN_TABLE)
                ky_name = ent_name & table_name
                
                Set tbl = catalog_.Tables(table_name)
                Set ky = tbl.Keys(ky_name)
                ' todo select by relatedcolumn = fky
                kycol = tbl.Keys(ky_name).Columns(0).Name
                kyfld = rs.Fields(kycol).Value
                
                qry1 = "SELECT * FROM [" & ent_name & "] WHERE " & fkey & " = " & kyfld
                Set tbls = FetchRecords((qry1), (ent_name), DbInfo, dbc, cnt + 1)
                propertyName = prop(EC_KEY_ENTITY_PROPERTY_NAME)
                
                If tbls Is Nothing Then
                Else
                    Call CallByName(T, propertyName, VbSet, tbls.Item(1))
                End If
            
            ElseIf prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
                
                fkey = prop(EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
                ent_name = prop(EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
                ky_name = table_name & ent_name
                
                Set tbl = catalog_.Tables(ent_name)
                kycol = tbl.Keys(ky_name).Columns(fkey).Name
                kyfld = rs.Fields(tbl.Keys(ky_name).Columns(fkey).RelatedColumn).Value
                
                qry1 = "SELECT * FROM [" & ent_name & "] WHERE " & fkey & " = " & kyfld
                Set tbls = FetchRecords((qry1), (ent_name), DbInfo, dbc, cnt + 1)
                propertyName = prop(EC_KEY_ENTITY_PROPERTY_NAME)
                    
                If tbls Is Nothing Then
                Else
                    Call CallByName(CallByName(T, propertyName, VbGet), "Items", VbSet, tbls)
                End If
            
            Else
                propertyName = prop(EC_KEY_ENTITY_PROPERTY_NAME)
                Call CallByName(T, propertyName, VbLet, rs.Fields(prop(EC_KEY_DB_FIELD_NAME)).Value)
            End If
        Next
        Items.Add T
        rs.MoveNext
    Loop
    rs.Close
    
    Set FetchRecords = Items
End Function

'Public Function FetchRecords(entity_name As String, DbInfo As Object, dbc As Object) As Collection
'    Dim rs As Recordset, Items As New Collection, field As ADODB.field, propertyName As String, T As Object
'    Dim table_name As String, props As Collection, prop As Collection, field_value
'
'    table_name = DbInfo(entity_name)(EC_KEY_DB_TABLE_NAME)
'    Set props = DbInfo(entity_name)(EC_KEY_ENTITY_PROPERTIES)
'
'    Set rs = FetchRecordSet(table_name)
'    rs.MoveFirst
'    Do Until rs.EOF
'        Set T = dbc.CreateObjectBy(entity_name)
'
''        For Each field In rs.Fields
'        For Each prop In props
'
'            ' todo add case of related values
'            If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Then
'                Dim fkey
'                fkey = prop(EC_KEY_ENTITY_FOREIGNKEY)
'
'                set tbls = FetchRecords(... , DBInfo, dbc)
'                Call CallByName(T, propertyName, VbLet, tbls)
'            Else
'                propertyName = prop(EC_KEY_ENTITY_PROPERTY_NAME)
'                Call CallByName(T, propertyName, VbLet, rs.Fields(prop(EC_KEY_DB_FIELD_NAME)).Value)
'            End If
'        Next
'        Items.Add T
'        rs.MoveNext
'    Loop
'    rs.Close
'
'    Set FetchRecords = Items
'End Function


Private Function GetPropertyNameByFieldName(props As Collection, field_name As String)
    Dim prop As Collection
    For Each prop In props
        If prop(EC_KEY_DB_FIELD_NAME) = field_name Then GetPropertyNameByFieldName = prop(EC_KEY_ENTITY_PROPERTY_NAME)
    Next
End Function

'Public Function FetchRecords(table_name As String, T As Object) As Collection
'    Dim rs As Recordset, Items As New Collection, field As ADODB.field, propertyName As String
'
'    Set rs = FetchRecordSet(table_name)
'    rs.MoveFirst
'    Do Until rs.EOF
'        For Each field In rs.Fields
'            propertyName = ConvertDbFieldNameToClassPropertyName(field.Name)
'            If IsNull(field.Value) Then
'                Call CallByName(T, propertyName, VbLet, 0)
'            Else
'                Call CallByName(T, propertyName, VbLet, field.Value)
'            End If
'        Next
'        Items.Add T
'        rs.MoveNext
'    Loop
'    rs.Close
'
'    Set FetchRecords = Items
'End Function

Public Function Insert(rec As Object, DbInfo As Object)
    Dim rs As New ADODB.Recordset, fl, vl
    Dim table_name As String, exceptional_field As String, exceptional_prop As Collection
        
    table_name = GetTableNameByEntity(rec, DbInfo)
    
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
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Insert()", Err.Description
End Function

Public Function Update(rec As Object, DbInfo As Object)
    Dim rs As New ADODB.Recordset, fl, vl
    
    fl = GetFieldsListByEntity(rec, DbInfo)
    vl = GetValuesListByEntity(rec, DbInfo)
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    Call rs.Open(Source:=GetTableNameByEntity(rec, DbInfo), ActiveConnection:=cn_, LockType:=adLockOptimistic)
        
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call MoveToRecord(rs, rec, DbInfo)
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
    
On Error GoTo errOpen
    If cn_.State <> adStateOpen Then cn_.Open
    Call rs.Open(Source:=GetTableNameByEntity(rec, DbInfo), ActiveConnection:=cn_, LockType:=adLockOptimistic)
        
    Call cn_.BeginTrans
On Error GoTo errTrans
    Call MoveToRecord(rs, rec, DbInfo)
    Call rs.Delete(adAffectCurrent)
    Call rs.Close
    Call cn_.CommitTrans
    Exit Function

errTrans:
    cn_.RollbackTrans
    If rs.State <> adStateClosed Then Call rs.Close
errOpen:
    Err.Raise Err.Number, Err.Source & vbCrLf & "EntityContexter.Delete()", Err.Description
End Function

Private Sub MoveToRecord(rs As ADODB.Recordset, rec As Object, DbInfo As Collection)
    Dim index_field As String, index_value, table_name As String, prop As Collection

    table_name = GetTableNameByEntity(rec, DbInfo)
    index_field = DbInfo(table_name)("PrimaryKey")
    For Each prop In DbInfo(table_name)(EC_KEY_ENTITY_PROPERTIES)
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
    Dim i As Long, value_list, table_name As String, props As Collection, prop As Collection, col As New Collection
        
    table_name = GetTableNameByEntity(rec, DbInfo)
    Set props = DbInfo(table_name)(EC_KEY_ENTITY_PROPERTIES)
    
    For Each prop In props
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_AUTOINCREMENT) = False Then col.Add CallByName(rec, prop(EC_KEY_ENTITY_PROPERTY_NAME), VbGet)
        End If
    Next
    GetValuesListByEntity = CollectionToArray(col)

End Function

Private Function SetAutoAutoincrementValues(rec As Object, DbInfo As Object, rs As ADODB.Recordset)
    Dim i As Long, table_name As String, prop As Collection
    
    table_name = GetTableNameByEntity(rec, DbInfo)
    i = 0
    For Each prop In DbInfo(table_name)(EC_KEY_ENTITY_PROPERTIES)
        i = i + 1
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_AUTOINCREMENT) = True Then
                Call CallByName(rec, prop(EC_KEY_ENTITY_PROPERTY_NAME), VbLet, rs.Fields(prop(EC_KEY_DB_FIELD_NAME)).Value)
            End If
        End If
    Next
End Function


Private Function GetFieldsListByEntity(rec As Object, DbInfo As Object)
    Dim i As Long, value_list, table_name As String, props As Collection, prop As Collection, col As New Collection
        
    table_name = GetTableNameByEntity(rec, DbInfo)
    Set props = DbInfo(table_name)(EC_KEY_ENTITY_PROPERTIES)
        
    For Each prop In props
        If prop(EC_KEY_ENTITY_EXIST_FOREIGNKEY) Or prop(EC_KEY_ENTITY_EXIST_INVERSE_FOREIGNKEY) Then
        Else
            If prop(EC_KEY_DB_AUTOINCREMENT) = False Then col.Add prop(EC_KEY_DB_FIELD_NAME)
        End If
    Next
    GetFieldsListByEntity = CollectionToArray(col)
        
End Function

Private Function GetTableNameByEntity(rec As Object, DbInfo As Object) As String
    Dim table_name As String, info As Collection
    
    For Each info In DbInfo
        If info(EC_KEY_ENTITY_NAME) = TypeName(rec) Then table_name = info(EC_KEY_DB_TABLE_NAME)
    Next
    
    GetTableNameByEntity = table_name
End Function

Private Function GetPrimaryKeyByEntity(rec As Object, DbInfo As Object) As String
    Dim info As Collection
    
    For Each info In DbInfo
        If info("EntityClassName") = TypeName(rec) Then GetPrimaryKeyByEntity = info("PrimaryKey")
    Next
    
End Function

Private Function GetDBInfoPropertyByFieldName(field_name As String, table_name As String, DbInfo As Object) As Collection
    Dim prop As Collection
    
    For Each prop In DbInfo(table_name)("Properties")
        If prop("DbFieldName") = field_name Then Set GetDBInfoPropertyByFieldName = prop
    Next
    
End Function

Private Property Get HasSchema() As Boolean
    HasSchema = schemaName_ <> ""
End Property



' Private Methods ----------------------------------------------------------------
Private Function FetchRecordSetByQuery(query_string As String) As Recordset
    If cn_.State <> adStateOpen Then cn_.Open
    Set FetchRecordSetByQuery = cn_.Execute(query_string)
End Function
Private Function FetchRecordSet(table_name As String) As Recordset
    Dim query As String
    
    If HasSchema Then
        query = "SELECT * FROM [" & schemaName_ & "].[" & table_name & "]"
    Else
        query = "SELECT * FROM [" & table_name & "]"
    End If
    
    If cn_.State <> adStateOpen Then cn_.Open
    Set FetchRecordSet = cn_.Execute(query)
End Function

Private Function GetTableNameFromDb() As Collection
    Dim tableNames As New Collection, sch As ADODB.Recordset
    
On Error GoTo errLabel
    If cn_.State <> adStateOpen Then cn_.Open
    Set sch = cn_.OpenSchema(adSchemaTables)
    
    sch.MoveFirst
    Do Until sch.EOF
        If HasSchema Then
            If sch.Fields("TABLE_SCHEMA").Value = schemaName_ Then Call tableNames.Add(Replace(sch.Fields("TABLE_NAME").Value, "$", ""))
        Else
            If sch.Fields("TABLE_TYPE").Value = "TABLE" Or sch.Fields("TABLE_TYPE").Value = "VIEW" Then Call tableNames.Add(Replace(sch.Fields("TABLE_NAME").Value, "$", ""))
        End If
        sch.MoveNext
    Loop
    cn_.Close
    Set GetTableNameFromDb = tableNames
    Exit Function

errLabel:
    Err.Raise 10000, Err.Source & vbCrLf & "at EntityContexter.GetTableNameFromDb()", Err.Description
End Function

Private Function GetEntityInfoByTableName(table_name As String) As Collection
    Dim props As New Collection, prop As Collection, tbl As ADOX.Table, col As ADOX.Column, entity_info As New Collection
    Dim ky As ADOX.Key

    Set tbl = catalog_.Tables(table_name)
    For Each col In tbl.Columns
        Set prop = GetPropertyInfoByTable(col)
        Call props.Add(prop, prop(EC_KEY_ENTITY_PROPERTY_NAME))
    Next
    
    Call entity_info.Add(props, EC_KEY_ENTITY_PROPERTIES)
    Call entity_info.Add(ConvertDbTableNameToEntityClassName(table_name), EC_KEY_ENTITY_NAME)
    Call entity_info.Add(ConvertDbTableNameToEntitesClassName(table_name), EC_KEY_ENTITIES_NAME)
    Call entity_info.Add(tbl, EC_KEY_ENTITY_DB_TABLE)
    Call entity_info.Add(table_name, EC_KEY_DB_TABLE_NAME)
    
    Set GetEntityInfoByTableName = entity_info
End Function

Private Function SetEntityInfoRelatedColumns(ByRef entityInfos As Collection)
    Dim EntityInfo As Collection, col As ADOX.Column, ky As ADOX.Key, tbl As ADOX.Table, prop As Collection
    
    For Each EntityInfo In entityInfos
        Set tbl = EntityInfo(EC_KEY_ENTITY_DB_TABLE)
        For Each ky In tbl.Keys
            If ky.Type <> adKeyForeign Then
            Else
                For Each col In ky.Columns
                    Set prop = GetPropertyInfoByForeignKey(ky.RelatedTable, col)
                    Call EntityInfo(EC_KEY_ENTITY_PROPERTIES).Add(prop, prop(EC_KEY_ENTITY_PROPERTY_NAME))
                    Set prop = GetPropertyInfoByInverseProperty(tbl.Name, col)
                    Call entityInfos(ky.RelatedTable).Item(EC_KEY_ENTITY_PROPERTIES).Add(prop, prop(EC_KEY_ENTITY_PROPERTY_NAME))
                Next
            End If
        Next
    Next
    
End Function

Private Function GetPropertyInfoByTable(col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String

    propertyName = ConvertDbFieldNameToClassPropertyName(col.Name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(col.Name)
    propertyTypeName = GetEntityPropertyTypeBy(col)

    Dim prop As New Collection
    Call prop.Add(propertyName, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(False, EC_KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col, EC_KEY_ENTITY_DB_COLUMN)

    Set GetPropertyInfoByTable = prop
End Function

Private Function GetPropertyInfoByForeignKey(key_table_name As String, col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertDbTableNameToEntityClassName(key_table_name)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(key_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(True, EC_KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add(col, EC_KEY_ENTITY_DB_COLUMN)
    Call prop.Add(col.RelatedColumn, EC_KEY_ENTITY_FOREIGNKEY)
    Call prop.Add(key_table_name, EC_KEY_ENTITY_FOREIGN_TABLE)
    Call prop.Add("", EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
    Call prop.Add("", EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    
    Set GetPropertyInfoByForeignKey = prop
End Function

Private Function GetPropertyInfoByInverseProperty(key_table_name As String, col As ADOX.Column)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
'    propertyName = ConvertDbTableNameToEntityClassName(key_table_name)
'    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(key_table_name)
'    propertyTypeName = ConvertDbTableNameToEntitesClassName(key_table_name)
    
    propertyName = ConvertDbTableNameToEntitesClassName(key_table_name)
    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(key_table_name)
    propertyTypeName = propertyName
    
    Dim prop As New Collection
    Call prop.Add(propertyName, EC_KEY_ENTITY_PROPERTY_NAME)
    Call prop.Add(propertyFieldName, EC_KEY_ENTITY_FIELD_NAME)
    Call prop.Add(propertyTypeName, EC_KEY_ENTITY_PROPERTY_TYPE)
    Call prop.Add(True, EC_KEY_ENTITY_IS_PROPERTY_RELATED)
    Call prop.Add("", EC_KEY_ENTITY_FOREIGNKEY)
    Call prop.Add("", EC_KEY_ENTITY_FOREIGN_TABLE)
    Call prop.Add(col.Name, EC_KEY_ENTITY_INVERSE_FOREIGNKEY)
    Call prop.Add(key_table_name, EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE)
    Call prop.Add(col, EC_KEY_ENTITY_DB_COLUMN)
    
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
    Dim texts As New Collection, str As String, txt
    Dim info As Collection, prop As Collection
    Const db_table_name As String = "DBTableName"

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
        .Add "Private query_ As String"
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
        .Add "Public Function Fetch(entity_name As String)"
        .Add ""
        .Add "    query_ = ""SELECT * FROM ["" & DbInfo(entity_name)(""DBTableName"") & ""]"" "
        .Add "    Set Fetch = ec_.FetchRecords(query_, entity_name, DbInfo, Me, 0)"
        .Add ""
        .Add "End Function"
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
'            .Add "    Set info = New Collection"
'            .Add "    Call info.Add(""" & info(EC_KEY_DB_TABLE_NAME) & """, """ & EC_KEY_DB_TABLE_NAME & """)"
'            .Add "    Call info.Add(""" & info(EC_KEY_ENTITY_NAME) & """, """ & EC_KEY_ENTITY_NAME & """)"
'            .Add "    Call info.Add(""" & info(EC_KEY_ENTITIES_NAME) & """, """ & EC_KEY_ENTITIES_NAME & """)"
            .Add "    Set info = ec_.CreateDbInfoBy(" & _
                        "  table_name   :=""" & info(EC_KEY_DB_TABLE_NAME) & """" & _
                        ", entity_name  :=""" & info(EC_KEY_ENTITY_NAME) & """" & _
                        ", entities_name:=""" & info(EC_KEY_ENTITIES_NAME) & """" & _
                        ")"
            If info(EC_KEY_ENTITY_DB_TABLE).Keys.Count > 0 Then .Add "    Call info.Add(""" & info(EC_KEY_ENTITY_DB_TABLE).Keys(0).Columns(0).Name & """, ""PrimaryKey"")"
            .Add ""
            .Add "    Set props = New Collection"
            For Each prop In info("Properties")
                If prop(EC_KEY_ENTITY_IS_PROPERTY_RELATED) = False Then
'                    .Add "    Set prop = New Collection"
'                    .Add "    Call prop.Add(""" & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """, """ & EC_KEY_ENTITY_PROPERTY_NAME & """)"
'                    .Add "    Call prop.Add(""" & prop(EC_KEY_ENTITY_FIELD_NAME) & """, """ & EC_KEY_ENTITY_FIELD_NAME & """)"
'                    .Add "    Call prop.Add(""" & prop(EC_KEY_ENTITY_PROPERTY_TYPE) & """, """ & EC_KEY_ENTITY_PROPERTY_TYPE & """)"
'                    .Add "    Call prop.Add(""" & prop(EC_KEY_ENTITY_DB_COLUMN).Name & """, """ & EC_KEY_DB_FIELD_NAME & """)"
'                    .Add "    Call prop.Add(""" & prop(EC_KEY_ENTITY_DB_COLUMN).Properties(EC_KEY_DB_AUTOINCREMENT) & """, """ & EC_KEY_DB_AUTOINCREMENT & """)"
                    
                    .Add "    Set prop = ec_.CreateDbInfoPropertyBy(" & _
                                    "property_name:=""" & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """" & _
                                    ", entity_field_name:=""" & prop(EC_KEY_ENTITY_FIELD_NAME) & """" & _
                                    ", entity_property_type:=""" & prop(EC_KEY_ENTITY_PROPERTY_TYPE) & """" & _
                                    ", db_field_name:=""" & prop(EC_KEY_ENTITY_DB_COLUMN).Name & """" & _
                                    ", auto_increment:=" & prop(EC_KEY_ENTITY_DB_COLUMN).Properties(EC_KEY_DB_AUTOINCREMENT) & _
                                    ")"
                    .Add "    Call props.Add(prop, """ & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """)"
                Else
'                    .Add "    Set prop = ec_.CreateDbInfoPropertyBy(" & _
'                                    "property_name:=""" & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """" & _
'                                    ", entity_field_name:=""" & prop(EC_KEY_ENTITY_FIELD_NAME) & """" & _
'                                    ", entity_property_type:=""" & prop(EC_KEY_ENTITY_PROPERTY_TYPE) & """" & _
'                                    ", foreign_key:=""" & prop(EC_KEY_ENTITY_FOREIGNKEY) & """" & _
'                                    ")"
                    If prop(EC_KEY_ENTITY_FOREIGNKEY) <> "" Then
                        .Add "    Set prop = ec_.CreateDbInfoPropertyBy(" & _
                                        "property_name:=""" & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """" & _
                                        ", foreign_key:=""" & prop(EC_KEY_ENTITY_FOREIGNKEY) & """" & _
                                        ", foreign_table:=""" & prop(EC_KEY_ENTITY_FOREIGN_TABLE) & """" & _
                                        ")"
                    Else
                        .Add "    Set prop = ec_.CreateDbInfoPropertyBy(" & _
                                        "property_name:=""" & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """" & _
                                        ", inverse_foreign_key:=""" & prop(EC_KEY_ENTITY_INVERSE_FOREIGNKEY) & """" & _
                                        ", inverse_foreign_table:=""" & prop(EC_KEY_ENTITY_INVERSE_FOREIGN_TABLE) & """" & _
                                        ")"
                    End If
                    .Add "    Call props.Add(prop , """ & prop(EC_KEY_ENTITY_PROPERTY_NAME) & """)"
                End If
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
    If col.Attributes = adColNullable Then
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
    ConvertDbTableNameToEntitesClassName = GetEstimatedPluralName(ConvertSnakeToPascalCase(dbTbleName))
End Function

Private Function ConvertDbTableNameToEntityClassName(dbTbleName As String) As String
    ConvertDbTableNameToEntityClassName = GetEstimatedSingularName(ConvertSnakeToPascalCase(dbTbleName))
End Function

Private Function ConvertDbTableNameToEntitiesClassFieldName(dbTbleName As String) As String
    ConvertDbTableNameToEntitiesClassFieldName = GetEstimatedPluralName(ConvertSnakeToCamelCase(dbTbleName)) & "_"
End Function

Private Function ConvertDbFieldNameToClassPropertyName(dbFieldName As String) As String
    ConvertDbFieldNameToClassPropertyName = ConvertSnakeToPascalCase(dbFieldName)
End Function

Private Function ConvertDBFieldNameToClassFieldName(dbFieldName As String) As String
    ConvertDBFieldNameToClassFieldName = ConvertSnakeToCamelCase(dbFieldName) & "_"
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


' Terminate ------------------------------------------------------------------------
Private Sub Class_Terminate()
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Sub

