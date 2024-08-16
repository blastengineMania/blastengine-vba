VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Client As Blastengine
Private m_DeliveryId As Long
Private m_From As From
Private m_To() As Dictionary
Private m_Cc() As String
Private m_Bcc() As String
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

Function Send(Optional ByVal dateReservationTime As Date = -1) As Boolean
    If UBound(m_To) = -1 Then
        Err.Raise 1, Description:="宛先は1つ以上指定してください"
    End If
    If dateReservationTime = -1 Then
        If UBound(m_To) > 1 Then
            If UBound(m_Cc) > 0 Or UBound(m_Bcc) > 0 Then
                Err.Raise 3, Description:="宛先が複数の場合、CCまたはBCCは指定できません"
            End If
            Send = SendBulk
        Else
            Send = SendTransaction
        End If
    Else
        If Not Not m_Cc Then
            If UBound(m_Cc) > 0 Then
                Err.Raise 2, Description:="予約配信は、CCを指定できません"
            End If
        End If
        If Not Not m_Bcc Then
            If UBound(m_Bcc) > 0 Then
                Err.Raise 2, Description:="予約配信は、BCCを指定できません"
            End If
        End If
        Send = SendBulk(dateReservationTime)
    End If
End Function

Function SendTransaction() As Boolean
    Dim i As Long
    Dim Mail As Transaction
    Set Mail = m_Client.Transaction
    Mail.Subject = m_Subject
    Mail.TextPart = m_TextPart
    Mail.Email = m_To(0)("email")
    If m_To(0).Exists("InsertCode") = True Then
        Mail.InsertCode = m_To(0)("InsertCode")
    End If
    Mail.HtmlPart = m_HtmlPart
    
    If Not Not m_Attachments Then
        If UBound(m_Attachments) > -1 Then
            For i = 0 To UBound(m_Attachments)
                Mail.Attachments(i) = m_Attachments(i)
            Next i
        End If
    End If
    If Not Not m_Cc Then
        If UBound(m_Cc) > -1 Then
            For i = 0 To UBound(m_Cc)
                Mail.Cc(i) = m_Cc(i)
            Next i
        End If
    End If
    If Not Not m_Bcc Then
        If UBound(m_Bcc) > -1 Then
            For i = 0 To UBound(m_Bcc)
                Mail.Bcc(i) = m_Bcc(i)
            Next i
        End If
    End If
    Mail.Encode = m_Encode
    Mail.From m_From.Email, m_From.Name
    Mail.ListUnsubscribe = m_ListUnsubscribe
    
    If Mail.Send Then
        SendTransaction = True
        m_DeliveryId = Mail.DeliveryId
    Else
        SendTransaction = False
        Set m_Error = Mail.Error
    End If
End Function

Function SendBulk(Optional ByVal dateReservationTime As Date = -1) As Boolean
    Dim i As Long
    Dim Mail As Bulk
    Set Mail = m_Client.Bulk
    Mail.Subject = m_Subject
    Mail.TextPart = m_TextPart
    If UBound(m_To) > -1 Then
        For i = 0 To UBound(m_To)
            Mail.Email m_To(i)("email"), m_To(i)("InsertCode")
        Next i
    End If
    Mail.HtmlPart = m_HtmlPart
    
    If Not Not m_Attachments Then
        If UBound(m_Attachments) > -1 Then
            For i = 0 To UBound(m_Attachments)
                Mail.Attachments(i) = m_Attachments(i)
            Next i
        End If
    End If
    Mail.Encode = m_Encode
    Mail.From m_From.Email, m_From.Name
    Mail.ListUnsubscribe = m_ListUnsubscribe
    If Mail.Begin = False Then
        Set m_Error = Mail.Error
        SendBulk = False
        Set Mail = Nothing
        Exit Function
    End If
    If UBound(m_To) > 50 Then
        If Mail.Import = False Then
            Set m_Error = Mail.Error
            SendBulk = False
            Set Mail = Nothing
            Exit Function
        End If
        While Mail.Job.Finished = False
            Application.Wait (Now + TimeValue("0:00:01"))
        Wend
        
        If Mail.Job.IsError Then
            Dim Val As Variant
            For Each Val In Mail.Job.Error
                m_Error.Add i, Val
            Next Val
            SendBulk = False
            Set Mail = Nothing
            Exit Function
        End If
    Else
        If Mail.Update = False Then
            Set m_Error = Mail.Error
            SendBulk = False
            Set Mail = Nothing
            Exit Function
        End If
    End If
    If Mail.Send Then
        SendBulk = True
        m_DeliveryId = Mail.DeliveryId
    Else
        SendBulk = False
        Set m_Error = Mail.Error
    End If
End Function


Function Delete() As Boolean
    Delete = m_Client.Delete(m_DeliveryId)
End Function

Property Let DeliveryId(ByVal Value As Long)
    m_DeliveryId = Value
End Property

Property Let Cc(i As Long, ByVal Email As String)
    ReDim Preserve m_Cc(i) As String
    m_Cc(i) = Email
End Property

Property Let Bcc(i As Long, ByVal Email As String)
    ReDim Preserve m_Bcc(i) As String
    m_Bcc(i) = Email
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

Sub Email(ByVal Email As String, Optional ByVal InsertCode As Dictionary = -1)
    On Error GoTo Skip
    Dim i As Long
    If (Not m_To) <> -1 Then
        i = UBound(m_To) + 1
    End If
    ReDim Preserve m_To(i) As Dictionary
    Dim Params As New Dictionary
    Params.Add "email", Email
    If InsertCode.Count <> -1 Then
        Params.Add "InsertCode", InsertCode
    End If
Skip:
    Set m_To(i) = Params
End Sub

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

