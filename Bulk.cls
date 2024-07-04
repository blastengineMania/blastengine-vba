VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bulk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Client As Blastengine
Private m_DeliveryId As Long
Private m_From As From
Private m_To() As Dictionary
Private m_Subject As String
Private m_ListUnsubscribe As ListUnsubscribe
Private m_Encode As String
Private m_TextPart As String
Private m_HtmlPart As String
Private m_Error As New Dictionary
Private m_Attachments() As String
Private m_Response As Dictionary

Private Sub Class_Initialize()
    m_Encode = "UTF-8"
End Sub

Property Let client(client As Blastengine)
    Set m_Client = client
End Property

Function Begin() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/bulk/begin"
    If (Not m_Attachments) = -1 Then
        Set dictResult = m_Client.SendText(HttpPost, strPath, New Dictionary, BuildParams)
    Else
        Set dictResult = m_Client.SendFiles(HttpPost, strPath, BuildParams, m_Attachments)
    End If
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Begin = False
    Else
        m_DeliveryId = dictResult("delivery_id")
        Begin = True
    End If
End Function

Function Update() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    ' Debug.Print JsonConverter.ConvertToJson(BuildParams)
    strPath = "/api/v1/deliveries/bulk/update/" & m_DeliveryId
    Set dictResult = m_Client.SendText(httpPut, strPath, New Dictionary, BuildParams)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Update = False
    Else
        ' m_DeliveryId = dictResult("delivery_id")
        Update = True
    End If
    Set dictResult = Nothing
End Function

Function Send(Optional ByVal dateReservationTime As Date = -1) As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    Dim dictBody As New Dictionary
    
    strPath = "/api/v1/deliveries/bulk/commit/" & m_DeliveryId
    If dateReservationTime = -1 Then
        strPath = strPath & "/immediate"
    Else
        Debug.Print m_Client.ConvertToISO8601(dateReservationTime)
        dictBody.Add "reservation_time", m_Client.ConvertToISO8601(dateReservationTime)
    End If
    Set dictResult = m_Client.SendText(httpPatch, strPath, New Dictionary, dictBody)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Send = False
    Else
        ' m_DeliveryId = dictResult("delivery_id")
        Send = True
    End If
End Function

Property Get DeliveryId() As Long
    DeliveryId = m_DeliveryId
End Property

Property Get Error() As Dictionary
    Set Error = m_Error
End Property

Sub Email(ByVal Email As String, Optional ByVal InsertCode As Dictionary = -1)
    Dim i As Long
    If (Not m_To) <> -1 Then
        i = UBound(m_To) + 1
    End If
    ReDim Preserve m_To(i) As Dictionary
    Dim Params As New Dictionary
    Params.Add "email", Email
    If InsertCode.Count > 0 Then
        Params.Add "insert_code", InsertCode
    End If
    Set m_To(i) = Params
End Sub

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


Function BuildParams() As Dictionary
    Set BuildParams = New Dictionary
    Dim From As Dictionary
    Set From = New Dictionary
    From.Add "email", m_From.Email
    If m_From.Name <> "" Then
        From.Add "name", m_From.Name
    End If
    BuildParams.Add "from", From
    BuildParams.Add "subject", m_Subject
    BuildParams.Add "text_part", m_TextPart
    If m_HtmlPart <> "" Then
        BuildParams.Add "html_part", m_HtmlPart
    End If
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

    If m_DeliveryId = 0 Then
        ' Only first time
        BuildParams.Add "encode", m_Encode
    Else
        ' Only update
        Dim Val As Variant
        Dim aryTo() As Dictionary
        ReDim aryTo(UBound(m_To)) As Dictionary
        Dim i As Long
        Dim Params As New Dictionary
        For Each Val In m_To
            Set Params = Val
            If Params.Exists("insert_code") Then
                Dim InsertCodes() As Dictionary
                InsertCodes = m_Client.InsertCodeToArray(Params("insert_code"))
                If (Not InsertCodes) <> -1 Then
                    Params("insert_code") = InsertCodes
                End If
            End If
            Set aryTo(i) = Params
            i = i + 1
        Next Val
        BuildParams.Add "to", aryTo
    End If
End Function

