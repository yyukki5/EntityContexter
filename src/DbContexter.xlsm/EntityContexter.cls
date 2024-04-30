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
' Version       : 0.1.0
' Project       : https://github.com/yyukki5/EntityContexter
' Depend on     : Microsoft ActiveX Data Objects X.X Library, Microsoft ADO Ext. 2.8 for DDL and Security
' Before Using  : Need to check "Trust access to the VBA project object model"
' Function      : Scaffolding DB
' ---------------------------------------------------------------------------------

Private Const VBA_CLASS_MODULE_ID As Long = 2

Private cn_ As ADODB.Connection
Private catalog_ As ADOX.catalog
Private isExportFiles_ As Boolean

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

Public Function Init(connectionString As String) As EntityContexter
    With cn_
        .connectionString = connectionString
        .Open
    End With

    catalog_.ActiveConnection = cn_.connectionString
    
    Set Init = Me
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

Public Function ScaffoldDataBese(Optional usingAttribute As Boolean = True) As Collection
    Dim tableName, tableNames As New Collection, entityInfos As New Collection, entityInfo As Collection
    
    ' setting
    isExportFiles_ = usingAttribute
    
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
'    CreateDbContect entityInfos
    Debug.Print "- Done -------"
    
    ' for test
    Set ScaffoldDataBese = entityInfos
End Function

Public Function GetKeyFieldNameByPropertyName(tbl As Object, propertyName As String) As String
    GetKeyFieldNameByPropertyName = CallByName(tbl, "GetKeyFieldName", VbMethod, propertyName)
End Function


' Private Methods ----------------------------------------------------------------
Private Function FetchRecordSet(tableName As String) As Recordset
    Dim query As String
    query = "SELECT * FROM [" & tableName & "]"
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
Private Function GetDbContextCode(tableEntityInfos As Collection) As String
    Dim texts As New Collection, str As String, txt
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

