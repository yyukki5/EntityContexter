VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private item_id_
Private visitor_id_
Private name_

Public Property Get ItemId(): ItemId = item_id_: End Property
Public Property Let ItemId(val): item_id_ = val: End Property
Public Property Get VisitorId(): VisitorId = visitor_id_: End Property
Public Property Let VisitorId(val): visitor_id_ = val: End Property
Public Property Get Name(): Name = name_: End Property
Public Property Let Name(val): name_ = val: End Property

Public Function GetNew() As KeyItem:    Set GetNew = New KeyItem: End Function
