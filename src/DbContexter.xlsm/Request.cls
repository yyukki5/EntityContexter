VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private request_id_
Private request_name_
Private requester_email_address_
Private create_at
Private visiters_ As New visiters

Public Property Get RequestId(): RequestId = request_id_: End Property
Public Property Let RequestId(val): request_id_ = val: End Property

Public Property Get RequesterName(): RequesterName = request_name_: End Property
Public Property Let RequesterName(val): request_name_ = val: End Property

Public Property Get RequesterEmailAddress(): RequesterEmailAddress = requester_email_address_: End Property
Public Property Let RequesterEmailAddress(val): requester_email_address_ = val: End Property

Public Property Get CreateAt(): CreateAt = create_at: End Property
Public Property Let CreateAt(val): create_at = val: End Property

Public Property Get visiters() As visiters: Set visiters = visiters_: End Property
Public Property Set visiters(val As visiters): Set visiters_ = val: End Property

Public Function GetNew() As Request:    Set GetNew = New Request: End Function


