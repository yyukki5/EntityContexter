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
' Project   : https://github.com/yyukki5/EntityContexter
' Depend on : Microsoft ActiveX Data Objects X.X Library
' Before Using : Need to check "Trust access to the VBA project object model"
' ---------------------------------------------------------------------------------

Private Const VBA_CLASS_MODULE_ID As Long = 2

Private cn_ As ADODB.Connection
Private catalog_ As ADOX.catalog
    
Private Enum PropertyNameTypeIndex
    [_Init]
    ecProperty
    ecField
    ecPropertyType
    ecPropertyIsLet
    ecKeyDbFieldName
    ecRelatedFieldName
    [_Last]
End Enum
Private Enum EntityTableNameIndex
    [_Init]
    ecTableName
    ecEntityName
    ecEntitesName
    ecPropertyNameType
    [_Last]
End Enum


Private Sub Class_Initialize()
    Set cn_ = New ADODB.Connection
    Set catalog_ = New ADOX.catalog
End Sub

Public Function Init() As EntityContexter
'    With cn_
'        .Provider = "Microsoft.ACE.OLEDB.12.0"
'        .Properties("Extended Properties") = "Excel 12.0"
'        .Open ThisWorkbook.Path & "\" & "TEST.xlsx"
'    End With

    With cn_
        .connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\resorce\db\Test.accdb"
        .Open
    End With

    catalog_.ActiveConnection = cn_.connectionString
    
    Set Init = Me
End Function

Function SetRelation(pk As Object, pKeyName As String, fk As Object, fKeyName As String)
    Dim pko As Object, fko As Object

    For Each pko In pk.items
        For Each fko In fk.items
            If CallByName(fko, fKeyName, VbGet) = CallByName(pko, pKeyName, VbGet) Then
                CallByName(pko, TypeName(fk), VbGet).Add fko
            End If
        Next
    Next
End Function

Public Function FetchRecords(tableName, T As Object) As Collection
    Dim rs As Recordset
    Dim items As New Collection
    Dim field As ADODB.field, fieldName As String
    
    Set rs = FetchRecordSet(CStr(tableName))
    rs.MoveFirst
    Do Until rs.EOF
        Set T = T.GetNew ' <- HACK more better. (to avoid GetNew() in each racord classes)
        For Each field In rs.Fields
            fieldName = Replace(StrConv(Replace(field.Name, "_", " "), vbProperCase), " ", "")
            Call CallByName(T, fieldName, VbLet, field.Value)
        Next
        items.Add T
        rs.MoveNext
    Loop
    rs.Close
    
    Set FetchRecords = items
End Function

Public Function ScaffoldDataBese()
    Dim tableName, tableNames As New Collection, tableEntityInfos As New Collection, entityInfo As Collection
    
    ' get table names
    Set tableNames = GetTableNameFromDb()
    
    ' create class module code from ADOX.catalog by table name
    For Each tableName In tableNames
        Set entityInfo = New Collection
        SetEntityInfoFromTable CStr(tableName), entityInfo
        tableEntityInfos.Add entityInfo, tableName
    Next
    ' Add relationship to class module code
    SetEntityInfoRelatedColumns tableEntityInfos
                
    ' create class module from table names
    For Each tableName In tableNames
        Debug.Print "- " & tableName & " -------"
        CreateEntityClassFromTableName tableEntityInfos(tableName)
        Debug.Print "----------------"
    Next
    
    ' create DbContext.cls
    'CreateDbontect()
End Function

Public Function GetKeyFieldNameByPropertyName(tbl As Object, propertyName As String) As String
    GetKeyFieldNameByPropertyName = CallByName(tbl, "GetKeyFieldName", VbMethod, propertyName)
End Function


' Private Methods ----------------------------------------------------------------
Private Function FetchRecordSet(tableName As String) As Recordset
    Dim query As String
    
    query = "SELECT * FROM [" & tableName & "$]"
    Set FetchRecordSet = cn_.Execute(query)
End Function

Private Function SetEntityInfoFromTable(tableName As String, ByRef entityInfo As Collection)
    Dim propertyNameType As New Collection, tbl As ADOX.Table, col As ADOX.Column, ky As ADOX.Key, dummy As Object, i As Long
    
    For Each col In catalog_.Tables(tableName).Columns
        propertyNameType.Add GetPropertyNameType(col.Name, col.Type)
    Next
    For i = EntityTableNameIndex.[_Init] To EntityTableNameIndex.[_Last]
        Select Case i
            Case ecTableName: entityInfo.Add tableName
            Case ecEntitesName: entityInfo.Add ConvertEntitiesClassName(tableName)
            Case ecEntityName: entityInfo.Add ConvertEntityClassName(tableName)
            Case ecPropertyNameType: entityInfo.Add propertyNameType
        End Select
    Next i
End Function
Private Function SetEntityInfoRelatedColumns(ByRef entityInfos As Collection)
    Dim entityInfo As Collection, propertyNameTypes As New Collection, col As ADOX.Column, ky As ADOX.Key, tableName As String
    
    For Each entityInfo In entityInfos
        tableName = entityInfo.item(ecTableName)
        For Each ky In catalog_.Tables(tableName).Keys
            If ky.Type <> adKeyForeign Then GoTo nextLoop
            For Each col In ky.Columns
                entityInfo.item(ecPropertyNameType).Add _
                    GetPropertyNameTypeByRelatedTableNto1(ky.RelatedTable, col.Name, col.RelatedColumn)
                entityInfos(ky.RelatedTable).item(ecPropertyNameType).Add _
                    GetPropertyNameTypeByRelatedTable1toN(tableName, col.RelatedColumn, col.Name)
            Next
nextLoop:
        Next
    Next
End Function

Private Function CreateEntityClassFromTableName(tableEntityInfo As Collection)
    Dim EntityClass As Object, EntitiesClass As Object, codeEntityClass As String, codeEntitiesClass As String
    
    codeEntityClass = GetEntityCode(tableEntityInfo(ecEntityName), tableEntityInfo(ecPropertyNameType))
    codeEntitiesClass = GetEntitesCode(tableEntityInfo(ecTableName), tableEntityInfo(ecEntityName))

    ' only force update (all delete and re-insert)
    ' EntityClass
    Set EntityClass = AddClassModule(tableEntityInfo(ecEntityName))
    EntityClass.CodeModule.DeleteLines 1, EntityClass.CodeModule.CountOfLines
    EntityClass.CodeModule.InsertLines 1, codeEntityClass
    Debug.Print tableEntityInfo(ecEntityName); ": done"
    
    Set EntitiesClass = AddClassModule(tableEntityInfo(ecEntitesName))
    EntitiesClass.CodeModule.DeleteLines 1, EntitiesClass.CodeModule.CountOfLines
    EntitiesClass.CodeModule.InsertLines 1, codeEntitiesClass
    Debug.Print tableEntityInfo(ecEntitesName); ": done"

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
Private Function GetEntitesCode(tableName As String, entityName As String) As String
    Dim texts As New Collection, str As String, txt
    
    With texts
        .Add "Option Explicit"
        .Add ""
        .Add "Private items_ As New Collection"
        .Add ""
        .Add "Public Property Get item(index As Long) As " & entityName
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
Private Function GetEntityCode(entityName As String, propertyNameType As Collection) As String
    Dim texts As New Collection, pNameType, str As String, txt
    
    With texts
        .Add "Option Explicit"
        .Add ""
        For Each pNameType In propertyNameType
            .Add "Private " & pNameType(ecField) & " as " & pNameType(ecPropertyType)
        Next
        .Add ""
        For Each pNameType In propertyNameType
            .Add "Public Property Get " & pNameType(ecProperty) & "() as " & pNameType(ecPropertyType)
            .Add "    " & pNameType(ecProperty) & " = " & pNameType(ecField)
            .Add "End Property"
            If pNameType(ecPropertyIsLet) Then
                .Add "Public Property Let " & pNameType(ecProperty) & "(val as " & pNameType(ecPropertyType) & ")"
                .Add "    Let " & pNameType(ecField) & " = val"
                .Add "End Property"
            Else
                .Add "Public Property Set " & pNameType(ecProperty) & "(val as " & pNameType(ecPropertyType) & ")"
                .Add "    Set " & pNameType(ecField) & " = val"
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
        For Each pNameType In propertyNameType
            .Add "        Case """ & pNameType(ecProperty) & """: s = """ & pNameType(ecKeyDbFieldName) & """"
        Next
        .Add "    End Select"
        .Add "    GetKeyFieldName = s"
        .Add "End Function"
    End With
    For Each txt In texts: str = str + txt + vbCrLf: Next
    
    GetEntityCode = str
End Function
Private Function GetPropertyNameType(fieldName As String, Optional fieldType As ADODB.DataTypeEnum, Optional actualValue As Variant)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
    
    propertyName = ConvertFieldNameToPropertyName(fieldName)
    propertyFieldName = ConvertFieldNameToClassFieldName(fieldName)
    If IsMissing(actualValue) = False Then propertyTypeName = TypeName(actualValue): GoTo gotoSetValue
    
    ' todo confirm
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
        Err.Raise 9100, "EntityContext.GetPropertyNameType()", "field type is not supported."
    End If

gotoSetValue:
    GetPropertyNameType = GetPropertyNameTypeArray(propertyFieldName, propertyName, propertyTypeName, True, fieldName, "")
End Function

Private Function GetPropertyNameTypeByRelatedTable1toN(keyTable, keyfield, relatedfieldName)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertEntitiesClassName(CStr(keyTable))
    propertyFieldName = ConvertToEntitiesClassNameField(CStr(keyTable))
    propertyTypeName = propertyName
    
    GetPropertyNameTypeByRelatedTable1toN = GetPropertyNameTypeArray(propertyFieldName, propertyName, propertyTypeName, False, keyfield, relatedfieldName)
End Function
Private Function GetPropertyNameTypeByRelatedTableNto1(keyTable, keyfield, relatedfieldName)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
                
    propertyName = ConvertEntityClassName(CStr(keyTable))
    propertyFieldName = ConvertFieldNameToClassFieldName(CStr(keyTable))
    propertyTypeName = propertyName
    
    GetPropertyNameTypeByRelatedTableNto1 = GetPropertyNameTypeArray(propertyFieldName, propertyName, propertyTypeName, False, keyfield, relatedfieldName)
End Function
Private Function GetPropertyNameTypeArray(propertyFieldName, propertyName, propertyTypeName, isLet, fieldName, relatedfieldName)
    Dim arr
    ReDim arr(PropertyNameTypeIndex.[_Init] + 1 To PropertyNameTypeIndex.[_Last] - 1)
    arr(PropertyNameTypeIndex.ecField) = propertyFieldName
    arr(PropertyNameTypeIndex.ecProperty) = propertyName
    arr(PropertyNameTypeIndex.ecPropertyType) = propertyTypeName
    arr(PropertyNameTypeIndex.ecPropertyIsLet) = isLet
    arr(PropertyNameTypeIndex.ecKeyDbFieldName) = fieldName
    arr(PropertyNameTypeIndex.ecRelatedFieldName) = relatedfieldName
    GetPropertyNameTypeArray = arr
End Function


Private Function ConvertEntitiesClassName(dbTbleName As String) As String
    ConvertEntitiesClassName = GetEstimatedPluralName(ConvertSnakeToPascalCase(dbTbleName))
End Function
Private Function ConvertEntityClassName(dbTbleName As String) As String
    ConvertEntityClassName = GetEstimatedSingularName(ConvertSnakeToPascalCase(dbTbleName))
End Function
Private Function ConvertToEntitiesClassNameField(dbTbleName As String) As String
    ConvertToEntitiesClassNameField = GetEstimatedPluralName(ConvertSnakeToCamelCase(dbTbleName)) & "_"
End Function
Private Function ConvertFieldNameToPropertyName(dbFieldName As String) As String
    ConvertFieldNameToPropertyName = ConvertSnakeToPascalCase(dbFieldName)
End Function
Private Function ConvertFieldNameToClassFieldName(dbFieldName As String) As String
    ConvertFieldNameToClassFieldName = ConvertSnakeToCamelCase(dbFieldName) & "_"
End Function
Private Function ConvertSnakeToPascalCase(snakeCase As String) As String
    ConvertSnakeToPascalCase = Replace(StrConv(Replace(snakeCase, "_", " "), vbProperCase), " ", "")
End Function
Private Function ConvertSnakeToCamelCase(snakeCase As String)
    Dim s As String
    s = Replace(StrConv(Replace(snakeCase, "_", " "), vbProperCase), " ", "")
    s = StrConv(Left(s, 1), vbLowerCase) & Mid(s, 2)
    ConvertSnakeToCamelCase = s
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
Private Function GetTableNameFromDb() As Collection
    Dim tableNames As New Collection, tbl As ADOX.Table
    
    For Each tbl In catalog_.Tables
        If tbl.Type = "TABLE" Then
            tableNames.Add Replace(tbl.Name, "$", "")
        End If
    Next

    Set GetTableNameFromDb = tableNames
End Function



' Terminate ------------------------------------------------------------------------
Private Sub Class_Terminate()
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Sub
