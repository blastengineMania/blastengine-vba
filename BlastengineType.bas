Attribute VB_Name = "BlastengineType"

Type From
    Email As String
    Name As String
End Type

Type InsertCode
    Key As String
    value As String
End Type

Type ListUnsubscribe
    Email As String
    Url As String
End Type

Enum StatusType
    EDIT = 1
    IMPORTING = 2
    RESERVE = 3
    WAIT = 4
    SENDING = 5
    SENT = 6
    FAILED = 7
End Enum

Enum DeliveryType
    Transaction = 1
    Bulk = 2
    SMTP = 3
End Enum

