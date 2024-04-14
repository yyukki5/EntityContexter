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
    ecProperty
    ecField
    ecPropertyType
    ecPropertyIsLet
End Enum

Private Sub Class_Initialize()
    Set cn_ = New ADODB.Connection
    Set catalog_ = New ADOX.catalog
End Sub

Public Function Init() As EntityContexter
    With cn_
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open ThisWorkbook.Path & "\" & "TEST.xlsx"
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
    Dim tableName, tableNames As New Collection
    
    ' get table names
    Set tableNames = GetTableNameFromDb()
    
    ' create class module from table names
    For Each tableName In tableNames
        CreateEntityClassFromTableName CStr(tableName)
    Next

    ' create DbContext.cls
    'CreateDbontect()
End Function




' Private Methods ----------------------------------------------------------------
Public Function FetchRecordSet(tableName As String) As Recordset
    Dim query As String
    
    query = "SELECT * FROM [" & tableName & "$]"
    Set FetchRecordSet = cn_.Execute(query)
End Function

' temp
Public Function CreateEntityClassFromTableName(tableName As String)
    
    Dim rs As ADODB.Recordset
    Dim propertyNameType As New Collection
    Dim entityName As String, entitesName As String, f As ADODB.field
    Dim EntityClass As Object, EntitiesClass As Object
        
    entitesName = Replace(StrConv(Replace(tableName, "_", " "), vbProperCase), " ", "")
    entityName = GetEstimatedSingularName(entitesName)
    
    Set rs = FetchRecordSet(tableName)
    For Each f In rs.Fields
        propertyNameType.Add GetPropertyNameType(f.Name, f.Type, f.Value)
    Next

    Set EntityClass = AddClassModule(entityName)
    EntityClass.CodeModule.DeleteLines 1, EntityClass.CodeModule.CountOfLines
    EntityClass.CodeModule.InsertLines 1, GetEntityCode(entityName, propertyNameType)
    
    Set EntitiesClass = AddClassModule(entitesName)
    EntitiesClass.CodeModule.DeleteLines 1, EntitiesClass.CodeModule.CountOfLines
    EntitiesClass.CodeModule.InsertLines 1, GetEntitesCode(entityName)
        
End Function

' temp
Public Function AddClassModule(className As String) As Object
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

Public Function GetEntitesCode(entityName As String) As String
    Dim texts As New Collection, str As String, txt
    
    With texts
        .Add "Option Explicit"
        .Add ""
        .Add "Private items_ As New Collection"
        .Add ""
        .Add "Public Property Get item(index As Long) As " & entityName & ": Set item = items_.item(index): End Property"
        .Add "Public Property Get items() As Collection: Set items = items_: End Property"
        .Add "Public Property Set items(items1 As Collection): Set items_ = items1: End Property"
        .Add ""
        .Add "Public Function Add(T As Object)"
        .Add "    items_.Add T"
        .Add "End Function"
    End With
    For Each txt In texts: str = str + txt + vbCrLf: Next
    
    GetEntitesCode = str
End Function

Public Function GetEntityCode(entityName As String, propertyNameType As Collection) As String
    Dim texts As New Collection, pNameType, str As String, txt
    
    With texts
        .Add "Option Explicit"
        .Add ""
        For Each pNameType In propertyNameType
            .Add "Private " & pNameType(ecField) & " as " & pNameType(ecPropertyType)
        Next
        .Add ""
        For Each pNameType In propertyNameType
            .Add "Public Property Get " & pNameType(ecProperty) & "() as " & pNameType(ecPropertyType) & ": " & pNameType(ecProperty) & " = " & pNameType(ecField) & ": End Property"
            If pNameType(ecPropertyIsLet) Then
                .Add "Public Property Let " & pNameType(ecProperty) & "(val as " & pNameType(ecPropertyType) & "): Let " & pNameType(ecField) & " = val: End Property"
            Else
                .Add "Public Property Set " & pNameType(ecProperty) & "(val as " & pNameType(ecPropertyType) & "): Set " & pNameType(ecField) & " = val: End Property"
            End If
        Next
        .Add "Public Function GetNew() As " & entityName & ": Set GetNew = New " & entityName & ": End Function"
    End With
    For Each txt In texts: str = str + txt + vbCrLf: Next
    
    GetEntityCode = str
End Function

'temp
Public Function GetPropertyNameType(fieldName As String, fieldType As ADODB.DataTypeEnum, Optional actualValue As Variant)
    Dim propertyName As String, propertyFieldName As String, propertyTypeName As String
    
    propertyName = Replace(StrConv(Replace(fieldName, "_", " "), vbProperCase), " ", "")
    propertyFieldName = fieldName & "_"
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
    GetPropertyNameType = Array(propertyName, propertyFieldName, propertyTypeName, True)
    
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

re:
    GetEstimatedSingularName = s
End Function

Private Function GetTableNameFromDb() As Collection
    Dim tableNames As New Collection, tbl
    
    For Each tbl In catalog_.Tables
        tableNames.Add Replace(tbl.Name, "$", "")
    Next

    Set GetTableNameFromDb = tableNames
End Function


' Terminate ------------------------------------------------------------------------
Private Sub Class_Terminate()
    If catalog_ Is Nothing Then Else Set catalog_ = Nothing
    If cn_ Is Nothing Then Else If cn_.State > 0 Then cn_.Close
End Sub
