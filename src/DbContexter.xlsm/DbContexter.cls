VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbContexter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cn_ As New ADODB.Connection

Public Function Init() As DbContexter
    With cn_
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open ThisWorkbook.Path & "\" & "TEST.xlsx"
    End With
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

Private Function FetchRecordSet(tableName As String) As Recordset
    Dim query As String
    
    query = "SELECT * FROM [" & tableName & "$]"
    Set FetchRecordSet = cn_.Execute(query)
End Function

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    If cn_ Is Nothing Then Else cn_.Close
End Sub
