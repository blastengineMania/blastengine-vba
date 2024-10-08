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
Private m_Job As Job

Private m_Status As String
Private m_DeliveryTime As Date
Private m_UpdatedTime As Date
Private m_CreatedTime As Date
Private m_ReservationTime As Date
Private m_DeliveryType As String
Private m_OpenCount As Long
Private m_TotalCount As Long
Private m_SentCount As Long
Private m_DropCount As Long
Private m_SoftErrorCount As Long
Private m_HardErrorCount As Long

Private Sub Class_Initialize()
    m_Encode = "UTF-8"
End Sub

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

Property Let DeliveryId(ByVal value As Long)
    m_DeliveryId = value
End Property


Function Begin() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/bulk/begin"
    If (Not m_Attachments) = -1 Then
        Set dictResult = m_Client.SendText(httpPost, strPath, New Dictionary, BuildParams)
    Else
        Set dictResult = m_Client.SendFiles(httpPost, strPath, BuildParams, m_Attachments)
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

Function Delete() As Boolean
    Delete = m_Client.Delete(m_DeliveryId)
End Function

Function Cancel() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/" & m_DeliveryId & "/cancel"
    Set dictResult = m_Client.SendText(httpPatch, strPath, New Dictionary, New Dictionary)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Cancel = False
    Else
        Cancel = True
    End If
    Set dictResult = Nothing
End Function


Function CreateCSV() As String
    Dim Headers() As String
    ReDim Headers(0) As String
    Headers(0) = "email"
    
    ' Check header
    Dim varTo As Variant
    Dim dicTo As Dictionary
    Dim dicInsertCode As Dictionary
    Dim varKey As Variant
    Dim Key As Variant
    Dim i As Long
    Dim aryLine() As String
    ReDim aryLine(UBound(m_To) + 1) As String
    
    
    For Each varTo In m_To
        Set dicTo = varTo
        If dicTo.Exists("InsertCode") Then
            Set dicInsertCode = dicTo.Item("InsertCode")
            For Each Key In dicInsertCode.Keys
                i = m_Client.IndexOf(Headers, Key)
                If i = -1 Then
                    ReDim Preserve Headers(UBound(Headers) + 1) As String
                    Headers(UBound(Headers)) = Key
                End If
            Next Key
        End If
    Next varTo
    
    aryLine(0) = Replace(Join(Headers, "__,__") & "__", "email__", "email")
    
    Dim lngRow As Long
    Dim aryRow() As String
    lngRow = 0
    For Each varTo In m_To
        lngRow = lngRow + 1
        Set dicTo = varTo
        ReDim aryRow(UBound(Headers) + 1) As String
        aryRow(0) = """" & dicTo("email") & """"
        i = 0
        If dicTo.Exists("InsertCode") = False Then
            aryLine(lngRow) = Join(aryRow, ",")
            GoTo Next_varTo
        End If
                
        Set dicInsertCode = dicTo("InsertCode")
        For Each Key In Headers
            If Key = "email" Then
                GoTo Next_Header
            End If
            i = i + 1
            If dicInsertCode.Exists(Key) Then
                aryRow(i) = """" & Replace(dicInsertCode(Key), """", """""") & """"
            Else
                aryRow(i) = ""
            End If
Next_Header:
         Next Key
         aryLine(lngRow) = Join(aryRow, ",")
Next_varTo:
    Next varTo
    CreateCSV = Join(aryLine, vbCrLf)
End Function

Function Import(Optional IgnoreErrors As Boolean = False, Optional Immediate As Boolean = False) As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    Dim aryFiles(0) As String
    
    Dim csvFile As String
    csvFile = m_Client.CreateTemporaryFile("csv")
    
    fileNum = FreeFile
    Open csvFile For Output As #fileNum
    Print #fileNum, CreateCSV()
    Close #fileNum
    
    aryFiles(0) = csvFile
    
    strPath = "/api/v1/deliveries/" & m_DeliveryId & "/emails/import"
    Dim Params As New Dictionary
    Params.Add "ignore_errors", IgnoreErrors
    Params.Add "immediate", Immediate
    Set dictResult = m_Client.SendFiles(httpPost, strPath, Params, aryFiles)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Import = False
    Else
        ' m_DeliveryId = dictResult("delivery_id")
        Import = True
        Set m_Job = New Job
        m_Job.Client = m_Client
        m_Job.Id = dictResult("job_id")
    End If
    Kill csvFile
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

Sub Email(ByVal MailAddress As String, Optional ByVal InsertCode As Dictionary = -1)
    On Error GoTo Skip
    Dim i As Long
    If (Not m_To) <> -1 Then
        i = UBound(m_To) + 1
    End If
    ReDim Preserve m_To(i) As Dictionary
    Dim Params As New Dictionary
    Params.Add "email", MailAddress
    If InsertCode.Count <> -1 Then
        Params.Add "InsertCode", InsertCode
    End If
Skip:
    Set m_To(i) = Params
End Sub

Property Get DeliveryId() As Long
    DeliveryId = m_DeliveryId
End Property

Property Get Error() As Dictionary
    Set Error = m_Error
End Property

Property Get Recipients() As Dictionary()
    Set Email = m_To
End Property

Property Get Job() As Job
    Set Job = m_Job
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
        If UBound(m_To) <= 50 Then
            ' Only update
            Dim Val As Variant
            Dim aryTo() As Dictionary
            ReDim aryTo(UBound(m_To)) As Dictionary
            Dim i As Long
            Dim Params As New Dictionary
            For Each Val In m_To
                Set Params = Val
                If Params.Exists("InsertCode") Then
                    Dim InsertCodes() As Dictionary
                    InsertCodes = m_Client.InsertCodeToArray(Params("InsertCode"))
                    If Not Not InsertCodes Then
                        If UBound(InsertCodes) > 0 Then
                            Params("insert_code") = InsertCodes
                        End If
                    End If
                End If
                Params.Remove "InsertCode"
                Set aryTo(i) = Params
                i = i + 1
            Next Val
            ' Debug.Print JsonConverter.ConvertToJson(aryTo)
            BuildParams.Add "to", aryTo
        End If
    End If
End Function

