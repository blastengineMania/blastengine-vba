VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Anchor As Long
Private m_Email As String
Private m_Count As Integer
Private m_DeliveryId As Long
Private m_DeliveryType() As String
Private m_Status() As String
Private m_ResponseCode() As String
Private m_DeliveryStart As Date
Private m_DeliveryEnd As Date

Private m_Client As Blastengine
Private m_Error As Dictionary

Private Sub Class_Initialize()
    Set m_Error = New Dictionary
    m_Count = 100
End Sub

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

Public Property Let Size(ByVal value As Integer)
    m_Size = value
End Property

Public Property Let Anchor(ByVal value As Integer)
    m_Anchor = value
End Property

Public Property Let Count(ByVal value As Integer)
    m_Count = value
End Property

Public Property Let DeliveryId(ByVal value As Long)
    m_DeliveryId = value
End Property

Public Property Let Email(ByVal value As String)
    m_Email = value
End Property

Public Property Let Status(i As Long, ByVal value As StatusType)
    ReDim Preserve m_Status(i) As String
    m_Status(i) = ToStatusString(value)
End Property

Public Property Let DeliveryType(i As Long, ByVal value As String)
    ReDim Preserve m_DeliveryType(i) As String
    m_DeliveryType(i) = value
End Property

Public Property Let ResponseCode(i As Long, ByVal value As String)
    ReDim Preserve m_ResponseCode(i) As String
    m_ResponseCode(i) = value
End Property

Function ToStatusString(value As StatusType) As String
    Select Case value
    Case StatusType.EDIT
        ToStatusString = "EDIT"
    Case StatusType.FAILED
        ToStatusString = "FAILED"
    Case StatusType.IMPORTING
        ToStatusString = "IMPORTING"
    Case StatusType.RESERVE
        ToStatusString = "RESERVE"
    Case StatusType.SENDING
        ToStatusString = "SENDING"
    Case StatusType.SENT
        ToStatusString = "SENT"
    Case StatusType.WAIT
        ToStatusString = "WAIT"
    End Select
End Function

Public Property Let DeliveryStart(ByVal value As Date)
    m_DeliveryStart = value
End Property

Public Property Let DeliveryEnd(ByVal value As Date)
    m_DeliveryEnd = value
End Property


Public Function Find() As Log()
    Dim strPath As String
    strPath = "/api/v1/logs/mails/results"
    Dim dictResult As Dictionary
    Set dictResult = m_Client.QueryRequest(strPath, BuildQuery())
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Info = False
        Exit Function
    End If
    Dim Results() As Log
    
    If dictResult.Exists("data") Then
        Dim Params As Variant
        Dim LogParams As Dictionary
        Dim i As Long
        i = 0
        For Each Params In dictResult("data")
            Dim Log As Log
            Set Log = m_Client.Log
            Set LogParams = Params
            
            Log.DeliveryId = CLng(LogParams("delivery_id"))
            Log.MaillogId = CLng(LogParams("maillog_id"))
            
            Log.DeliveryType = LogParams("delivery_type")
            Log.Email = LogParams("email")
            Log.Status = LogParams("status")
            Log.LastResponseCode = CInt(LogParams("last_response_code"))
            If LogParams.Exists("delivery_time") And Not IsNull(LogParams("delivery_time")) Then
                Log.DeliveryTime = m_Client.ConvertISO8601ToDate(LogParams("delivery_time"))
            End If
            If LogParams.Exists("open_time") And Not IsNull(LogParams("open_time")) Then
                Log.OpenTime = m_Client.ConvertISO8601ToDate(LogParams("open_time"))
            End If
            If LogParams.Exists("created_time") And Not IsNull(LogParams("created_time")) Then
                Log.CreatedTime = m_Client.ConvertISO8601ToDate(LogParams("created_time"))
            End If
            If LogParams.Exists("updated_time") And Not IsNull(LogParams("updated_time")) Then
                Log.DeliveryTime = m_Client.ConvertISO8601ToDate(LogParams("updated_time"))
            End If
            
            ReDim Preserve Results(i) As Log
            Set Results(i) = Log
            i = i + 1
        Next Params
    End If
    Find = Results
    ' Debug.Print JsonConverter.ConvertToJson(dictResult)
End Function


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    m_Client.AddLongQuery BuildQuery, "anchor", m_Anchor
    m_Client.AddStringQuery BuildQuery, "email", m_Email
    If m_DeliveryStart <> Empty Then
        BuildQuery.Add "delivery_start", m_Client.ConvertToISO8601(m_DeliveryStart)
    End If
    If m_DeliveryEnd <> Empty Then
        BuildQuery.Add "delivery_end", m_Client.ConvertToISO8601(m_DeliveryEnd)
    End If
    m_Client.AddIntegerQuery BuildQuery, "count", m_Count
    m_Client.AddLongQuery BuildQuery, "delivery_id", m_DeliveryId
    
    m_Client.ArrayToString BuildQuery, "status", m_Status
    m_Client.ArrayToString BuildQuery, "delivery_type", m_DeliveryType
    m_Client.ArrayToString BuildQuery, "response_code", m_ResponseCode

End Function

