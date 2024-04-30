Attribute VB_Name = "Sample"
Option Explicit


Sub Sample()
    Dim dbc As New DbContextSample
    Dim tbl1s As New Table1s
    
    dbc.Init
    Set tbl1s = dbc.Table1s
    Debug.Print tbl1s(1).Table2.Gen         ' 2
    Debug.Print tbl1s(1).Table2.MemberName  ' subaru.oozora
    Debug.Print tbl1s(1).Table3s(1).TagName ' protein the subaru
    Debug.Print tbl1s(1).Ddate              ' 2018/09/16

End Sub

