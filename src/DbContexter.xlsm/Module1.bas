Attribute VB_Name = "Module1"
Option Explicit


Sub Sample()

    Dim dbc As New DbContext
    dbc.Init
    
'    reqs = dbc.Requests.WhereEvaluate("x => x.Name = 'abc'").Any()
    
    Dim v As visiter, r As Request
    For Each r In dbc.Requests.items
        For Each v In r.visiters.items
            If v.KeyItems.items.Count > 0 Then
                Debug.Print r.RequesterName & ", " & vbTab & v.VisiterId & ":" & v.Name & " --- " & v.KeyItems.item(1).Name
            Else
                Debug.Print r.RequesterName & ", " & vbTab & v.VisiterId & ":" & v.Name
            End If
        Next
    Next
End Sub


