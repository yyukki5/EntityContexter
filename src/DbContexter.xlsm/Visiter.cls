VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Visiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private visiter_id_
Private request_id_
Private name_
Private key_items_ As New KeyItems

Public Property Get VisiterId(): VisiterId = visiter_id_: End Property
Public Property Let VisiterId(val): visiter_id_ = val: End Property
Public Property Get RequestId(): RequestId = request_id_: End Property
Public Property Let RequestId(val): request_id_ = val: End Property
Public Property Get Name(): Name = name_: End Property
Public Property Let Name(val): name_ = val: End Property
Public Property Get KeyItems() As KeyItems: Set KeyItems = key_items_: End Property
Public Property Set KeyItems(val As KeyItems): Set key_items_ = val: End Property
Public Function GetNew() As visiter:    Set GetNew = New visiter: End Function

Public Function Init() As visiter
    

End Function

