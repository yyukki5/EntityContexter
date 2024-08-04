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

Private cn_ As ADODB.Connection
Private catalog_ As ADOX.catalog
Private isExportFiles_ As Boolean
Private schemaName_ As String
Private connectionString_ As String

Private Enum PropertyInfoIndex
    [_Init]
    ecPropertyName
    ecFieldName
    ecPropertyType
    ecPropertyIsLet
    ecDbKeyFieldName
    ecDbRelatedFieldName
    [_Last]
End Enum
Private Enum EntityInfoIndex
    [_Init]
    ecDbTableName
    ecEntityClassName
    ecEntitesClassName
    ecPropertyInfo
    [_Last]
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

Function SetRelationShip1toN(primaryKey As Object, primaryKeyPropertyName As String, foreignKey As Object, foreignKeyPropertyName As String)
    Dim primaryKeyObject As Object, foreignKeyObject As Object

    For Each primaryKeyObject In primaryKey.items
        For Each foreignKeyObject In foreignKey.items
            If CallByName(foreignKeyObject, foreignKeyPropertyName, VbGet) = CallByName(primaryKeyObject, primaryKeyPropertyName, VbGet) Then
                CallByName(primaryKeyObject, TypeName(foreignKey), VbGet).Add foreignKeyObject
                Call CallByName(foreignKeyObject, TypeName(primaryKeyObject), VbSet, primaryKeyObject)
            End If
        Next
    Next
End Function

Function SetRelationShipNto1(primaryKey As Object, primaryKeyPropertyName As String, foreignKey As Object, foreignKeyPropertyName As String)
    Dim primaryKeyObject As Object, foreignKeyObject As Object

    For Each primaryKeyObject In primaryKey.items
        For Each foreignKeyObject In foreignKey.items
            If CallByName(foreignKeyObject, foreignKeyPropertyName, VbGet) = CallByName(primaryKeyObject, primaryKeyPropertyName, VbGet) Then
                Call CallByName(primaryKeyObject, TypeName(foreignKey.item(1)), VbSet, foreignKeyObject)
                Call CallByName(foreignKeyObject, TypeName(primaryKey), VbSet, primaryKey)
            End If
        Next
    Next
End Function

Public Function FetchRecords(tableName, T As Object) As Collection
    Dim rs As Recordset
    Dim items As New Collection
    Dim field As ADODB.field, propertyName As String
    
    Set rs = FetchRecordSet(CStr(tableName))
    rs.MoveFirst
    Do Until rs.EOF
        Set T = T.GetNew ' <- HACK more better. (to avoid GetNew() in each racord classes)
        For Each field In rs.Fields
            propertyName = ConvertDbFieldNameToClassPropertyName(field.Name)
            If IsNull(field.Value) Then
                Call CallByName(T, propertyName, VbLet, 0)
            Else
                Call CallByName(T, propertyName, VbLet, field.Value)
            End If
        Next
        items.Add T
        rs.MoveNext
    Loop
    rs.Close
    
    Set FetchRecords = items
End Function


Public Function ScaffoldDataBese(Optional usingAttribute As Boolean = True, Optional schema_name As String) As Collection
    Dim tableName, tableNames As New Collection, entityInfos As New Collection, entityInfo As Collection
    
    ' setting
    isExportFiles_ = usingAttribute
    schemaName_ = schema_name
    
    ' get table names
    Set tableNames = GetTableNameFromDb()
    
    ' create class module code from ADOX.catalog by table name
    For Each tableName In tableNames
        Set entityInfo = New Collection
        SetEntityInfoFromTable CStr(tableName), entityInfo
        entityInfos.Add entityInfo, tableName
    Next
    ' Add relationship to class module code
    SetEntityInfoRelatedColumns entityInfos
                
    Debug.Print "- " & tableName & " -------"
    ' create class module from table names
    For Each tableName In tableNames
        CreateEntityClassFromTableName entityInfos(tableName)
        Debug.Print tableName & ": done"
    Next
    ' to be implement:  create DbContext.cls
    Call CreateDbContext(entityInfos, dbcontext_name)
    Debug.Print "- Done -------"
    
    ' for test
    Set ScaffoldDataBese = entityInfos
End Function
Private Sub MoveToRecord(rs As ADODB.Recordset, rec As Object, DbInfo As Collection)
    Dim index_field As String, index_value, table_name As String, prop As Collection

Public Function GetKeyFieldNameByPropertyName(tbl As Object, propertyName As String) As String
    GetKeyFieldNameByPropertyName = CallByName(tbl, "GetKeyFieldName", VbMethod, propertyName)

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
    For Each prop In DbInfo(table_name)("Properties")
        If prop("dbFieldName") = index_field Then index_value = CallByName(rec, prop("PropertyName"), VbGet)
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
    Set props = DbInfo(table_name)("Properties")
    
    For Each prop In props
        If prop("Autoincrement") = False Then _
            col.Add CallByName(rec, prop("PropertyName"), VbGet)
    Next
    GetValuesListByEntity = CollectionToArray(col)

End Function

Private Function SetAutoAutoincrementValues(rec As Object, DbInfo As Object, rs As ADODB.Recordset)
    Dim i As Long, table_name As String, prop As Collection
    
    table_name = GetTableNameByEntity(rec, DbInfo)
    i = 0
    For Each prop In DbInfo(table_name)("Properties")
        i = i + 1
        If prop("Autoincrement") = True Then
            Call CallByName(rec, prop("PropertyName"), VbLet, rs.Fields(prop("DbFieldName")).Value)
        End If
    Next
End Function


Private Function GetFieldsListByEntity(rec As Object, DbInfo As Object)
    Dim i As Long, value_list, table_name As String, props As Collection, prop As Collection, col As New Collection
        
    table_name = GetTableNameByEntity(rec, DbInfo)
    Set props = DbInfo(table_name)("Properties")
        
    For Each prop In props
        If prop("Autoincrement") = False Then _
            col.Add prop("dbFieldName")
    Next
    GetFieldsListByEntity = CollectionToArray(col)
        
End Function

Private Function GetTableNameByEntity(rec As Object, DbInfo As Object) As String
    Dim table_name As String, info As Collection
    
    For Each info In DbInfo
        If info("EntityClassName") = TypeName(rec) Then table_name = info("DBTableName")
    Next
    
    GetTableNameByEntity = table_name
End Function

Private Function GetPrimaryKeyByEntity(rec As Object, DbInfo As Object) As String
    Dim info As Collection
    
    For Each info In DbInfo
        If info("EntityClassName") = TypeName(rec) Then GetPrimaryKeyByEntity = info("PrimaryKey")
    Next
    
End Function

Private Function GetDBInfoPropertyByFieldName(field_Name As String, table_name As String, DbInfo As Object) As Collection
    Dim prop As Collection
    
    For Each prop In DbInfo(table_name)("Properties")
        If prop("DbFieldName") = field_Name Then Set GetDBInfoPropertyByFieldName = prop
    Next
    
End Function

Private Property Get HasSchema() As Boolean
    HasSchema = schemaName_ <> ""
End Property



' Private Methods ----------------------------------------------------------------
Private Function FetchRecordSet(tableName As String) As Recordset
    Dim query As String
    
    If HasSchema Then
        query = "SELECT * FROM [" & schemaName_ & "].[" & tableName & "]"
    Else
    query = "SELECT * FROM [" & tableName & "]"
    End If
    
    If cn_.State <> adStateOpen Then cn_.Open
    Set FetchRecordSet = cn_.Execute(query)
End Function

Private Function GetTableNameFromDb() As Collection
    Dim tableNames As New Collection, tbl As ADOX.Table
    
    For Each tbl In catalog_.Tables
        If tbl.Type = "TABLE" Then tableNames.Add Replace(tbl.Name, "$", "")
    Next

    Set GetTableNameFromDb = tableNames
End Function
Private Function SetEntityInfoFromTable(tableName As String, ByRef entityInfo As Collection)
    Dim propertyNameType As New Collection, tbl As ADOX.Table, col As ADOX.Column, ky As ADOX.Key, dummy As Object, i As Long
    
    For Each col In catalog_.Tables(tableName).Columns
        propertyNameType.Add GetPropertyInfoAsTable(col.Name, col.Type)
    Next
    For i = EntityInfoIndex.[_Init] To EntityInfoIndex.[_Last]
        Select Case i
            Case ecDbTableName: entityInfo.Add tableName
            Case ecEntitesClassName: entityInfo.Add ConvertDbTableNameToEntitesClassName(tableName)
            Case ecEntityClassName: entityInfo.Add ConvertDbTableNameToEntityClassName(tableName)
            Case ecPropertyInfo: entityInfo.Add propertyNameType
        End Select
    Next i
End Function
Private Function SetEntityInfoRelatedColumns(ByRef entityInfos As Collection)
    Dim entityInfo As Collection, propertyNameTypes As New Collection, col As ADOX.Column, ky As ADOX.Key, tableName As String
    
    For Each entityInfo In entityInfos
        tableName = entityInfo.item(ecDbTableName)
        For Each ky In catalog_.Tables(tableName).Keys
            If ky.Type <> adKeyForeign Then GoTo nextLoop
            For Each col In ky.Columns
                entityInfo.item(ecPropertyInfo).Add _
                    GetPropertyInfoAsRelatedTableNto1(ky.RelatedTable, col.Name, col.RelatedColumn)
                entityInfos(ky.RelatedTable).item(ecPropertyInfo).Add _
                    GetPropertyInfoAsRelatedTable1toN(tableName, col.RelatedColumn, col.Name)
            Next
nextLoop:
        Next
    Next
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
        
    codeEntityClass = GetEntityCode(tableEntityInfo(ecEntityClassName), tableEntityInfo(ecPropertyInfo))
    codeEntitiesClass = GetEntitesCode(tableEntityInfo(ecEntitesClassName), tableEntityInfo(ecEntityClassName))

    If isExportFiles_ Then
        ' option1 using write file and import file. to using attribute key word in class module.
        AddClassModuleFromFile tableEntityInfo(EntityInfoIndex.ecEntityClassName), codeEntityClass
        AddClassModuleFromFile tableEntityInfo(EntityInfoIndex.ecEntitesClassName), codeEntitiesClass
    Else
        ' option2 using CodeModule.Insertlines()
        ' only force update (all delete and re-insert)
        Set EntityClass = AddClassModule(tableEntityInfo(ecEntityClassName))
        EntityClass.CodeModule.DeleteLines 1, EntityClass.CodeModule.CountOfLines
        EntityClass.CodeModule.InsertLines 1, codeEntityClass
        Debug.Print tableEntityInfo(ecEntityClassName); ": done"
        
        Set EntitiesClass = AddClassModule(tableEntityInfo(ecEntitesClassName))
        EntitiesClass.CodeModule.DeleteLines 1, EntitiesClass.CodeModule.CountOfLines
        EntitiesClass.CodeModule.InsertLines 1, codeEntitiesClass
    End If
    
End Function

' to be implement
'Private Function CreateDbContect(entityInfos As Collection)
'End Function
'Private Function GetDbcCode(entityInfos As Collection)
'End Function

Private Function AddClassModuleFromFile(className, codes)
    Dim classFilePath As String
    classFilePath = ThisWorkbook.path & "\" & className & ".cls"
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
    Const DB_TABLE_NAME As String = "DBTableName"
    
    
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
            .Add "    Set info = New Collection"
            .Add "    Call info.Add(""" & info("DBTableName") & """, ""DBTableName"")"
            .Add "    Call info.Add(""" & info("EntityClassName") & """, ""EntityClassName"")"
            .Add "    Call info.Add(""" & info("EntitiesClassName") & """, ""EntitiesClassName"")"
            .Add "    Call info.Add(""" & info("PrimaryKey") & """, ""PrimaryKey"")"
            .Add ""
            .Add "    Set props = New Collection"
            For Each prop In info("Properties")
                If prop("IsLet") Then
                    .Add "    Set prop = New Collection"
                    .Add "    Call prop.Add(""" & prop("PropertyName") & """, ""PropertyName"")"
                    .Add "    Call prop.Add(""" & prop("ClassFieldName") & """, ""ClassFieldName"")"
                    .Add "    Call prop.Add(""" & prop("PropertyType") & """, ""PropertyType"")"
                    .Add "    Call prop.Add(""" & prop("DbFieldName") & """, ""DbFieldName"")"
                    .Add "    Call prop.Add(""" & prop("ForeignKey") & """, ""ForeignKey"")"
                    .Add "    Call prop.Add(" & prop("Autoincrement") & ", ""Autoincrement"")"
                    '.Add "    Call prop.Add(""" & prop(4) & """, ""PrimaryKey"")"
                    .Add "    Call props.Add(prop, """ & prop("PropertyName") & """)"
                End If
            Next
            .Add "    Call info.Add(props, ""Properties"")"
            .Add "    Call infos.Add(info, """ & info("DBTableName") & """)"
            .Add ""
            .Add ""
        Next
        .Add ""
        .Add "    Set DbInfo = infos"
        .Add "End Property"
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
        .Add "Public Property Get item(index As Long) As " & entityName
        If isExportFiles_ Then
            .Add "Attribute item.VB_UserMemId = 0"
        End If
        .Add "    Set item = items_.item(index)"
        .Add "End Property"
        .Add "Public Property Get items() As Collection"
        .Add "    Set items = items_"
        .Add "End Property"
        .Add "Public Property Set items(items1 As Collection)"
        .Add "    Set items_ = items1"
        .Add "End Property"
        .Add ""
        .Add "Public Function Add(T As Object)"
        .Add "    items_.Add T"
        .Add "End Function"

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
            If pInfo(ecPropertyIsLet) Then
                .Add "Private " & pInfo(ecFieldName) & " as " & pInfo(ecPropertyType)
            Else
                .Add "Private " & pInfo(ecFieldName) & " as New " & pInfo(ecPropertyType)
            End If
        Next
        .Add ""
        For Each pInfo In propertyInfo
            If pInfo(ecPropertyIsLet) Then
                .Add "Public Property Get " & pInfo(ecPropertyName) & "() as " & pInfo(ecPropertyType)
                .Add "    Let " & pInfo(ecPropertyName) & " = " & pInfo(ecFieldName)
                .Add "End Property"
                .Add "Public Property Let " & pInfo(ecPropertyName) & "(val as " & pInfo(ecPropertyType) & ")"
                .Add "    Let " & pInfo(ecFieldName) & " = val"
                .Add "End Property"
            Else
                .Add "Public Property Get " & pInfo(ecPropertyName) & "() as " & pInfo(ecPropertyType)
                .Add "    Set " & pInfo(ecPropertyName) & " = " & pInfo(ecFieldName)
                .Add "End Property"
                .Add "Public Property Set " & pInfo(ecPropertyName) & "(val as " & pInfo(ecPropertyType) & ")"
                .Add "    Set " & pInfo(ecFieldName) & " = val"
                .Add "End Property"
            End If
        Next
        .Add ""
        .Add "Public Function GetNew() As " & entityName
        .Add "    Set GetNew = New " & entityName
        .Add "End Function"
        .Add "Public Function GetKeyFieldName(propertyName As String) As String"
        .Add "    Dim s As String"
        .Add "    Select Case propertyName"
        For Each pInfo In propertyInfo
            .Add "        Case """ & pInfo(ecPropertyName) & """: s = """ & pInfo(ecDbKeyFieldName) & """"
        Next
        .Add "    End Select"
        .Add "    GetKeyFieldName = s"
        .Add "End Function"
    End With
    For Each txt In texts: str = str + txt + vbCrLf: Next
    
    GetEntityCode = str
End Function
Private Function GetPropertyInfoAsTable(fieldName As String, Optional fieldType As ADODB.DataTypeEnum, Optional actualValue As Variant)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
    
    propertyName = ConvertDbFieldNameToClassPropertyName(fieldName)
    propertyFieldName = ConvertDBFieldNameToClassFieldName(fieldName)
    If IsMissing(actualValue) = False Then propertyTypeName = TypeName(actualValue): GoTo gotoSetValue
    
    If fieldType = adInteger Then
        propertyTypeName = "Integer"
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
    Else
        Err.Raise 9100, "EntityContext.GetPropertyInfoAsTable()", "field type is not supported."
    End If

gotoSetValue:
    GetPropertyInfoAsTable = GetPropertyInfo(propertyFieldName, propertyName, propertyTypeName, True, fieldName, "")
End Function
Private Function GetPropertyInfoAsRelatedTable1toN(keyTable, keyfield, relatedfieldName)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertDbTableNameToEntitesClassName(CStr(keyTable))
    propertyFieldName = ConvertDbTableNameToEntitiesClassFieldName(CStr(keyTable))
    propertyTypeName = propertyName
    
    GetPropertyInfoAsRelatedTable1toN = GetPropertyInfo(propertyFieldName, propertyName, propertyTypeName, False, keyfield, relatedfieldName)
End Function
Private Function GetPropertyInfoAsRelatedTableNto1(keyTable, keyfield, relatedfieldName)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertDbTableNameToEntityClassName(CStr(keyTable))
    propertyFieldName = ConvertDBFieldNameToClassFieldName(CStr(keyTable))
    propertyTypeName = propertyName
    
    GetPropertyInfoAsRelatedTableNto1 = GetPropertyInfo(propertyFieldName, propertyName, propertyTypeName, False, keyfield, relatedfieldName)
End Function
Private Function GetPropertyInfo(propertyFieldName, propertyName, propertyTypeName, isLet, fieldName, relatedfieldName)
    Dim arr
    ReDim arr(PropertyInfoIndex.[_Init] + 1 To PropertyInfoIndex.[_Last] - 1)
    arr(PropertyInfoIndex.ecFieldName) = propertyFieldName
    arr(PropertyInfoIndex.ecPropertyName) = propertyName
    arr(PropertyInfoIndex.ecPropertyType) = propertyTypeName
    arr(PropertyInfoIndex.ecPropertyIsLet) = isLet
    arr(PropertyInfoIndex.ecDbKeyFieldName) = fieldName
    arr(PropertyInfoIndex.ecDbRelatedFieldName) = relatedfieldName
    GetPropertyInfo = arr
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
    ConvertSnakeToPascalCase = Replace(StrConv(Replace(snakeCase, "_", " "), vbProperCase), " ", "")
End Function
Private Function ConvertSnakeToCamelCase(snakeCase As String)
    Dim s As String
    s = Replace(StrConv(Replace(snakeCase, "_", " "), vbProperCase), " ", "")
    ConvertSnakeToCamelCase = StrConv(Left(s, 1), vbLowerCase) & Mid(s, 2)
End Function
Private Function GetEstimatedPluralName(singularName As String) As String
    Dim s As String
    
    If Right(singularName, 2) = "sh" Then _
        s = Left(singularName, Len(singularName) - 2) & "shes": GoTo re
    If Right(singularName, 2) = "ch" Then _
        s = Left(singularName, Len(singularName) - 2) & "ches": GoTo re
    If Right(singularName, 1) = "y" Then _
        s = Left(singularName, Len(singularName) - 1) & "ies": GoTo re
    If Right(singularName, 1) = "s" Then _
        s = Left(singularName, Len(singularName) - 1) & "ses": GoTo re
    If Right(singularName, 1) = "x" Then _
        s = Left(singularName, Len(singularName) - 2) & "xes": GoTo re
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


' Terminate ------------------------------------------------------------------------
Private Sub Class_Terminate()
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Sub

