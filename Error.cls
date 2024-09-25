VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ErrorStart As Date
Private m_ErrorEnd As Date
Private m_Email As String
Private m_ResponseCode() As Integer

Private m_Client As Blastengine
Private m_Error As Dictionary
Private m_Job As Job

Private Sub Class_Initialize()
    Set m_Error = New Dictionary
    m_Size = 100
    m_Page = 1
    m_Sort = "delivery_time:desc"
End Sub

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

Public Property Let ErrorStart(ByVal value As Date)
    m_ErrorStart = value
End Property

Public Property Let ErrorEnd(ByVal value As Date)
    m_ErrorEnd = value
End Property

Public Property Let Email(ByVal value As String)
   m_Email = value
End Property

Public Property Get Error() As Dictionary
    Set Error = m_Error
End Property

Public Property Get Job() As Job
    Set Job = m_Job
End Property

Public Property Let ResponseCode(i As Long, ByVal value As Integer)
    ReDim Preserve m_ResponseCode(i) As Integer
    m_ResponseCode(i) = value
End Property


Public Function Start() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/errors/list"
    Set dictResult = m_Client.SendText(httpPost, strPath, New Dictionary, BuildQuery)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Start = False
    Else
        ' m_DeliveryId = dictResult("delivery_id")
        Import = True
        Set m_Job = New Job
        m_Job.Client = m_Client
        m_Job.Id = dictResult("job_id")
        m_Job.JobType = "Error"
        Start = True
    End If
    Set dictResult = Nothing
End Function

Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    m_Client.AddStringQuery BuildQuery, "email", m_Email
    If m_ErrorStart <> Empty Then
        BuildQuery.Add "error_start", m_Client.ConvertToISO8601(m_ErrorStart)
    End If
    If m_ErrorEnd <> Empty Then
        BuildQuery.Add "error_end", m_Client.ConvertToISO8601(m_ErrorEnd)
    End If
    If (Not m_ResponseCode) <> -1 Then
        BuildQuery.Add "response_code", m_ResponseCode
    End If
End Function

