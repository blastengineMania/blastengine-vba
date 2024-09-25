VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_DeliveryTime As Date
Private m_DeliveryId As Long
Private m_MaillogId As Long
Private m_DeliveryType As String
Private m_Email As String
Private m_Status As String
Private m_LastResponseCode As Integer
Private m_LastResponseMessage As String
Private m_OpenTime As Date
Private m_CreatedTime As Date
Private m_UpdatedTime As Date

Private m_History() As History

Private m_Client As Blastengine
Private m_Error As Dictionary

Private Sub Class_Initialize()
    Set m_Error = New Dictionary
End Sub

Function Histories() As History()
    
    If Not Not m_History Then
        Histories = m_History
        Exit Function
    End If
    Dim strPath As String
    strPath = "/api/v1/logs/mails/" & m_MaillogId
    Dim dictResult As Dictionary
    Set dictResult = m_Client.QueryRequest(strPath, New Dictionary)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Exit Function
    End If
    Dim Results() As History
    
    If dictResult.Exists("sent_history") Then
        Dim Params As Variant
        Dim HistoryParams As Dictionary
        Dim i As Long
        i = 0
        For Each Params In dictResult("sent_history")
            Dim History As History
            Set HistoryParams = Params
            Set History = New History
            
            History.Status = m_Client.StringToStatus(HistoryParams("status"))
            
            History.ResponseCode = CInt(HistoryParams("response_code"))
            If HistoryParams.Exists("delivery_time") And Not IsNull(HistoryParams("delivery_time")) Then
                History.DeliveryTime = m_Client.ConvertISO8601ToDate(HistoryParams("delivery_time"))
            End If
            History.ResponseMessage = HistoryParams("response_message")
            
            ReDim Preserve Results(i) As History
            Set Results(i) = History
            i = i + 1
        Next Params
    End If
    m_History = Results
    Histories = m_History
    ' Debug.Print JsonConverter.ConvertToJson(dictResult)
End Function

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

' m_DeliveryTime ?????
Public Property Get DeliveryTime() As Date
    DeliveryTime = m_DeliveryTime
End Property

Public Property Let DeliveryTime(ByVal value As Date)
    m_DeliveryTime = value
End Property

' m_DeliveryId ?????
Public Property Get DeliveryId() As Long
    DeliveryId = m_DeliveryId
End Property

Public Property Let DeliveryId(ByVal value As Long)
    m_DeliveryId = value
End Property

' m_MaillogId ?????
Public Property Get MaillogId() As Long
    MaillogId = m_MaillogId
End Property

Public Property Let MaillogId(ByVal value As Long)
    m_MaillogId = value
End Property

' m_DeliveryType ?????
Public Property Get DeliveryType() As String
    DeliveryType = m_DeliveryType
End Property

Public Property Let DeliveryType(ByVal value As String)
    m_DeliveryType = value
End Property

' m_Email ?????
Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Email(ByVal value As String)
    m_Email = value
End Property

' m_Status ?????
Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Status(ByVal value As String)
    m_Status = value
End Property

' m_LastResponseCode ?????
Public Property Get LastResponseCode() As Integer
    LastResponseCode = m_LastResponseCode
End Property

Public Property Let LastResponseCode(ByVal value As Integer)
    m_LastResponseCode = value
End Property

' m_LastResponseMessage ?????
Public Property Get LastResponseMessage() As String
    LastResponseMessage = m_LastResponseMessage
End Property

Public Property Let LastResponseMessage(ByVal value As String)
    m_LastResponseMessage = value
End Property

' m_OpenTime ?????
Public Property Get OpenTime() As Date
    OpenTime = m_OpenTime
End Property

Public Property Let OpenTime(ByVal value As Date)
    m_OpenTime = value
End Property

' m_CreatedTime ?????
Public Property Get CreatedTime() As Date
    CreatedTime = m_CreatedTime
End Property

Public Property Let CreatedTime(ByVal value As Date)
    m_CreatedTime = value
End Property

' m_UpdatedTime ?????
Public Property Get UpdatedTime() As Date
    UpdatedTime = m_UpdatedTime
End Property

Public Property Let UpdatedTime(ByVal value As Date)
    m_UpdatedTime = value
End Property


