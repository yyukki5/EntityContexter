VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Visiters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private items_ As New Collection

Public Property Get item(index As Long) As visiter
    Set item = items_.item(index)
End Property

Public Property Get items() As Collection
    Set items = items_
End Property
Public Property Set items(items1 As Collection)
    Set items_ = items1
End Property

Public Function Add(T As Object)
    items_.Add T
End Function
