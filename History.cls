VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Status As StatusType
Private m_DeliveryTime As Date
Private m_ResponseCode As Integer
Private m_ResponseMessage As String

Private Sub Class_Initialize()
End Sub

' m_Status ????? (StatusType???)
Public Property Get Status() As StatusType
    Status = m_Status
End Property

Public Property Let Status(ByVal value As StatusType)
    m_Status = value
End Property

' m_DeliveryTime ?????
Public Property Get DeliveryTime() As Date
    DeliveryTime = m_DeliveryTime
End Property

Public Property Let DeliveryTime(ByVal value As Date)
    m_DeliveryTime = value
End Property

' m_ResponseCode ?????
Public Property Get ResponseCode() As Integer
    ResponseCode = m_ResponseCode
End Property

Public Property Let ResponseCode(ByVal value As Integer)
    m_ResponseCode = value
End Property

' m_ResponseMessage ?????
Public Property Get ResponseMessage() As String
    ResponseMessage = m_ResponseMessage
End Property

Public Property Let ResponseMessage(ByVal value As String)
    m_ResponseMessage = value
End Property

