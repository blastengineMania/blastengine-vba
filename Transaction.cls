VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Client As Blastengine
Private m_DeliveryId As Long
Private m_From As From
Private m_To As String
Private m_Cc() As String
Private m_Bcc() As String
Private m_InsertCode As Dictionary
Private m_Subject As String
Private m_ListUnsubscribe As ListUnsubscribe
Private m_Encode As String
Private m_TextPart As String
Private m_HtmlPart As String
Private m_Error As Dictionary
Private m_Attachments() As String
Private m_Response As Dictionary

Private m_DeliveryType As String
Private m_Status As String
Private m_TotalCount As Long
Private m_SentCount As Long
Private m_DropCount As Long
Private m_HardErrorCount As Long
Private m_SoftErrorCount As Long
Private m_OpenCount As Long
Private m_DeliveryTime As Date
Private m_ReservationTime As Date
Private m_CreatedTime As Date
Private m_UpdatedTime As Date

Private Sub Class_Initialize()
    m_Encode = "UTF-8"
    Set m_InsertCode = New Dictionary
    Set m_Error = New Dictionary
End Sub

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

Function Info() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/" & m_DeliveryId
    Set dictResult = m_Client.SendText(httpGet, strPath, New Dictionary, New Dictionary)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Info = False
        Exit Function
    End If
    m_From.Name = dictResult("from")("name") & ""
    m_From.Email = dictResult("from")("email")
    m_Subject = dictResult("subject")
    m_TextPart = dictResult("text_part")
    m_HtmlPart = dictResult("html_part") & ""
    
    m_Client.SetDate m_DeliveryTime, dictResult("delivery_time")
    m_Client.SetDate m_ReservationTime, dictResult("reservation_time")
    m_Client.SetDate m_CreatedTime, dictResult("created_time")
    m_Client.SetDate m_UpdatedTime, dictResult("updated_time")
    
    m_DeliveryType = dictResult("delivery_type")
    m_Status = dictResult("status")
    m_TotalCount = dictResult("total_count")
    m_SentCount = dictResult("sent_count")
    m_DropCount = dictResult("drop_count")
    m_HardErrorCount = dictResult("hard_error_count")
    m_SoftErrorCount = dictResult("soft_error_count")
    m_OpenCount = dictResult("open_count")
    
    ' Debug.Print JsonConverter.ConvertToJson(dictResult)
    Info = True
End Function

Function Send() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/transaction"
    If (Not m_Attachments) = -1 Then
        Set dictResult = m_Client.SendText(httpPost, strPath, New Dictionary, BuildParams)
    Else
        Set dictResult = m_Client.SendFiles(httpPost, strPath, BuildParams, m_Attachments)
    End If
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Send = False
    Else
        m_DeliveryId = dictResult("delivery_id")
        Send = True
    End If
End Function

Function Delete() As Boolean
    Delete = m_Client.Delete(m_DeliveryId)
End Function

Property Let DeliveryId(ByVal Value As Long)
    m_DeliveryId = Value
End Property

Property Let Email(ByVal Email As String)
    m_To = Email
End Property

Property Let Cc(i As Long, ByVal Email As String)
    ReDim Preserve m_Cc(i) As String
    m_Cc(i) = Email
End Property

Property Let Bcc(i As Long, ByVal Email As String)
    ReDim Preserve m_Bcc(i) As String
    m_Bcc(i) = Email
End Property

Property Get DeliveryId() As Long
    DeliveryId = m_DeliveryId
End Property

Property Get Status() As String
    Status = m_Status
End Property

Property Get DeliveryTime() As Date
    DeliveryTime = m_DeliveryTime
End Property

Property Get UpdatedTime() As Date
    UpdatedTime = m_UpdatedTime
End Property

Property Get CreatedTime() As Date
    CreatedTime = m_CreatedTime
End Property

Property Get ReservationTime() As Date
    ReservationTime = m_ReservationTime
End Property

Property Get DeliveryType() As String
    DeliveryType = m_DeliveryType
End Property

Property Get OpenCount() As Long
    OpenCount = m_OpenCount
End Property
Property Get TotalCount() As Long
    TotalCount = m_TotalCount
End Property
Property Get SentCount() As Long
    SentCount = m_SentCount
End Property
Property Get DropCount() As Long
    DropCount = m_DropCount
End Property
Property Get SoftErrorCount() As Long
    SoftErrorCount = m_SoftErrorCount
End Property
Property Get HardErrorCount() As Long
    HardErrorCount = m_HardErrorCount
End Property


Property Get Error() As Dictionary
    Set Error = m_Error
End Property

Sub From(Email As String, Optional Name As String = "")
    Dim From As From
    From.Email = Email
    From.Name = Name
    m_From = From
End Sub

Property Let Attachments(i As Long, ByVal FilePath As String)
    ReDim Preserve m_Attachments(i) As String
    m_Attachments(i) = FilePath
End Property

Property Let Subject(ByVal Subject As String)
    m_Subject = Subject
End Property

Property Let Encode(ByVal Encode As String)
    m_Encode = Encode
End Property

Property Let TextPart(ByVal TextPart As String)
    m_TextPart = TextPart
End Property

Property Let HtmlPart(ByVal HtmlPart As String)
    m_HtmlPart = HtmlPart
End Property

Property Let ListUnsubscribe(ListUnsubscribe As ListUnsubscribe)
    m_ListUnsubscribe = ListUnsubscribe
End Property

Property Let InsertCode(ByVal InsertCode As Dictionary)
    Set m_InsertCode = InsertCode
End Property

Function BuildParams() As Dictionary
    Set BuildParams = New Dictionary
    Dim From As Dictionary
    Set From = New Dictionary
    From.Add "email", m_From.Email
    If m_From.Name <> "" Then
        From.Add "name", m_From.Name
    End If
    BuildParams.Add "from", From
    BuildParams.Add "to", m_To
    If (Not m_Cc) <> -1 Then
        BuildParams.Add "cc", m_Cc
    End If
    If (Not m_Bcc) <> -1 Then
        BuildParams.Add "bcc", m_Bcc
    End If
    Dim Key As Variant
    Dim i As Long
    i = 0
    Dim InsertCodes() As Dictionary
    InsertCodes = m_Client.InsertCodeToArray(m_InsertCode)
    If (Not InsertCodes) > -1 Then
        BuildParams.Add "insert_code", InsertCodes
    End If
    BuildParams.Add "subject", m_Subject
    BuildParams.Add "text_part", m_TextPart
    If m_HtmlPart <> "" Then
        BuildParams.Add "html_part", m_HtmlPart
    End If
    BuildParams.Add "encode", m_Encode
    
    If m_ListUnsubscribe.Email <> "" Or m_ListUnsubscribe.Url <> "" Then
        Dim ListUnsubscribe As Dictionary
        Set ListUnsubscribe = New Dictionary
        If m_ListUnsubscribe.Email <> "" Then
            ListUnsubscribe.Add "mailto", "mailto:" & m_ListUnsubscribe.Email
        End If
        If m_ListUnsubscribe.Url <> "" Then
            ListUnsubscribe.Add "url", "mailto:" & m_ListUnsubscribe.Url
        End If
        BuildParams.Add "list_unsubscribe", ListUnsubscribe
    End If

End Function
