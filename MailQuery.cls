VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_TextPart As String
Private m_HtmlPart As String
Private m_Subject As String
Private m_From As String
Private m_ListUnsubscribeMailto As String
Private m_ListUnsubscribeUrl As String
Private m_Status() As String
Private m_DeliveryType() As String
Private m_DeliveryStart As Date
Private m_DeliveryEnd As Date
Private m_Size As Integer
Private m_Page As Integer
Private m_Sort As String
Private m_Client As Blastengine
Private m_Error As Dictionary

Private Sub Class_Initialize()
    Set m_Error = New Dictionary
    m_Size = 100
    m_Page = 1
    m_Sort = "delivery_time:desc"
End Sub

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

Public Property Let Status(i As Long, ByVal value As StatusType)
    ReDim Preserve m_Status(i) As String
    m_Status(i) = ToStatusString(value)
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

Public Property Let DeliveryType(i As Long, ByVal value As String)
    ReDim Preserve m_DeliveryType(i) As String
    m_DeliveryType(i) = value
End Property

Public Property Let TextPart(ByVal value As String)
    m_TextPart = value
End Property

Public Property Let HtmlPart(ByVal value As String)
    m_HtmlPart = value
End Property

Public Property Let Subject(ByVal value As String)
    m_Subject = value
End Property

Public Property Let From(ByVal value As String)
    m_From = value
End Property

Public Property Let ListUnsubscribeMailto(ByVal value As String)
    m_ListUnsubscribeMailto = value
End Property

Public Property Let ListUnsubscribeUrl(ByVal value As String)
    m_ListUnsubscribeUrl = value
End Property

Public Property Let Sort(ByVal value As String)
    m_Sort = value
End Property

Public Property Let DeliveryStart(ByVal value As Date)
    m_DeliveryStart = value
End Property

Public Property Let DeliveryEnd(ByVal value As Date)
    m_DeliveryEnd = value
End Property

Public Property Let Size(ByVal value As Integer)
    m_Size = value
End Property

Public Property Let Page(ByVal value As Integer)
    m_Page = value
End Property

Public Function Find() As Mail()
    Find = Search("/api/v1/deliveries")
End Function

Public Function All() As Mail()
    All = Search("/api/v1/deliveries/all")
End Function

Public Function Search(strPath As String) As Mail()
    Dim dictResult As Dictionary
    Set dictResult = m_Client.QueryRequest(strPath, BuildQuery())
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Info = False
        Exit Function
    End If
    Dim Results() As Mail
    
    If dictResult.Exists("data") Then
        Dim Params As Variant
        Dim MailParams As Dictionary
        Dim i As Long
        i = 0
        For Each Params In dictResult("data")
            Dim Mail As Mail
            Set Mail = m_Client.Mail
            Dim MailFrom As Dictionary
            Set MailParams = Params
            Mail.UpdatedTime = m_Client.ConvertISO8601ToDate(MailParams("updated_time"))
            Mail.CreatedTime = m_Client.ConvertISO8601ToDate(MailParams("created_time"))
            Mail.DeliveryType = MailParams("delivery_type")
            Mail.Subject = MailParams("subject")
            Set MailFrom = MailParams("from")
            Mail.From MailFrom("email"), IIf(IsNull(MailFrom("name")), "", MailFrom("name"))
            If MailParams.Exists("reservation_time") And Not IsNull(MailParams("reservation_time")) Then
                Mail.ReservationTime = m_Client.ConvertISO8601ToDate(MailParams("reservation_time"))
            End If
            If MailParams.Exists("delivery_time") And Not IsNull(MailParams("delivery_time")) Then
                Mail.DeliveryTime = m_Client.ConvertISO8601ToDate(MailParams("delivery_time"))
            End If
            Mail.Status = MailParams("status")
            
            ReDim Preserve Results(i) As Mail
            Set Results(i) = Mail
            i = i + 1
        Next Params
    End If
    Search = Results
    ' Debug.Print JsonConverter.ConvertToJson(dictResult)
End Function


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    m_Client.AddStringQuery BuildQuery, "text_part", m_TextPart
    m_Client.AddStringQuery BuildQuery, "html_part", m_HtmlPart
    m_Client.AddStringQuery BuildQuery, "subject", m_Subject
    m_Client.AddStringQuery BuildQuery, "from", m_From
    m_Client.AddStringQuery BuildQuery, "list_unsubscribe_mailto", m_ListUnsubscribeMailto
    m_Client.AddStringQuery BuildQuery, "list_unsubscribe_url", m_ListUnsubscribeUrl
    m_Client.AddStringQuery BuildQuery, "sort", m_Sort
    If m_DeliveryStart <> Empty Then
        BuildQuery.Add "delivery_start", m_Client.ConvertToISO8601(m_DeliveryStart)
    End If
    If m_DeliveryEnd <> Empty Then
        BuildQuery.Add "delivery_end", m_Client.ConvertToISO8601(m_DeliveryEnd)
    End If
    m_Client.AddIntegerQuery BuildQuery, "size", m_Size
    m_Client.AddIntegerQuery BuildQuery, "page", m_Page
    
    m_Client.ArrayToString BuildQuery, "status", m_Status
    m_Client.ArrayToString BuildQuery, "delivery_type", m_DeliveryType

End Function
