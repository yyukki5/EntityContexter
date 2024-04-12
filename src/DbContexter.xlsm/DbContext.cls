VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Depend on :  Active X Data ... *.* ~

'Need information by User
'   DB info
'   Table Name & Record Class & Records Class  <- To be automation by DbContexter.cls
'   Relation Table + KeyName

' conversion
' field name ->  fieldname+"_"  private field name
' field name ->  replace _ to "" and to Pascal case  : property name



'table (record + "s")
Public Requests As New Requests
Public Visitors As New visiters
Public KeyItems As New KeyItems

Sub Init()
    Dim dbc As New DbContexter
    ' record
    Dim req As New Request
    Dim visitor As New visiter
    Dim kyitem As New KeyItem
    
    dbc.Init
    Set Requests.items = dbc.FetchRecords("Sheet1", req)
    Set Visitors.items = dbc.FetchRecords("Sheet2", visitor)
    Set KeyItems.items = dbc.FetchRecords("Sheet3", kyitem)
    
    ' Relations
    dbc.SetRelation Visitors, "VisiterId", KeyItems, "VisitorId"
    dbc.SetRelation Requests, "RequestId", Visitors, "RequestId"
End Sub

