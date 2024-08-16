Attribute VB_Name = "Module1"
Private Declare PtrSafe Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringW" ( _
    ByVal pbBinary As LongPtr, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As LongPtr, _
    ByVal pcchString As LongPtr _
    ) As Long

Private Const CRYPT_STRING_BASE64 As Long = &H1&

'�o�C�g�z���BASE64�ŃG���R�[�h����Unicode������ɂ���֐�
'���̊֐��Ƀo�C�g�z���n���ۂɁAUTF-8�ϊ������o�C�g�z���n���̂�
'SJIS�ϊ������o�C�g�z���n���̂��AUTF16�ϊ�(���ϊ�)�����o�C�g�z���n���̂���
'���R�Ȃ���o�͂���镶���񂪕ς���Ă��܂��B
Public Function CryptBytesToString(ByRef bData() As Byte) As String
    Dim pbBinary As LongPtr
    Dim cbBinary As Long
    If UBound(bData) = -1 Then
        Exit Function
    End If

    '�o�C�g�z��̐擪�A�h���X�@�Ɓ@�������擾
    pbBinary = VarPtr(bData(0))
    cbBinary = UBound(bData) - LBound(bData) + 1

    Dim nBufferSize As Long
    '�ϊ��㕶����̒������擾���A�ϊ��㕶������i�[���邾���̃o�b�t�@��p��
    If CryptBinaryToString(pbBinary, cbBinary, CRYPT_STRING_BASE64, _
                                0, VarPtr(nBufferSize)) Then
        CryptBytesToString = String(nBufferSize, vbNullChar)
        '�K�v�Ȓ������p�ӂ����󕶎���o�b�t�@�̐擪�A�h���X(StrPtr)����ϊ����ʂŏ㏑��
        If CryptBinaryToString(pbBinary, cbBinary, CRYPT_STRING_BASE64, _
                                StrPtr(CryptBytesToString), VarPtr(nBufferSize)) Then
        End If
    End If
End Function


