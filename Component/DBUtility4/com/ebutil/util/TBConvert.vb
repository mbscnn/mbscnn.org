Imports TBConvertCLR

Public Class TBConvert

    '���J��X��  UCS2-->BIG5
    Private Const m_sUCS2BIG5FilePath As String = "\TB_UCS2_BIG5.bin"
    Private Shared m_oUCS2BIG5TB As TBConvertCLR.TB_UCS2_BIG5 = New TBConvertCLR.TB_UCS2_BIG5

    '���J��X��  UTF-->BIG5
    Private Const m_sUTF2BIG5FilePath As String = m_sUCS2BIG5FilePath
    Private Shared m_oUTF2BIG5TB As TBConvertCLR.TB_UCS2_BIG5 = New TBConvertCLR.TB_UCS2_BIG5

    'TBConvertUTF8BIG5

    Private Shared m_TBCLR As System.Collections.Hashtable = init()
    Private Shared m_TBStatus As System.Collections.Hashtable

    Shared Function init() As System.Collections.Hashtable
        m_TBCLR = New System.Collections.Hashtable

        '���J��X��  UCS2-->BIG5
        If TBConvertCLR.UCS2Convert.TBConvertLoadUCS2BIG5(m_sUCS2BIG5FilePath, m_oUCS2BIG5TB) = TBConvertCLR.TBConvertReturnVal.CONVT_SUCCESS Then
            m_TBCLR.Add(m_sUCS2BIG5FilePath, m_oUCS2BIG5TB)
            m_TBStatus.Add(m_sUCS2BIG5FilePath, TBConvertCLR.TBConvertReturnVal.CONVT_SUCCESS)
        Else
            m_TBStatus.Add(m_sUCS2BIG5FilePath, TBConvertCLR.TBConvertReturnVal.CONVT_LOAD_TABLE_FAIL)
        End If

        '���J��X��  UTF-->BIG5
        If TBConvertCLR.UCS2Convert.TBConvertLoadUCS2BIG5(m_sUTF2BIG5FilePath, m_oUTF2BIG5TB) = TBConvertCLR.TBConvertReturnVal.CONVT_SUCCESS Then
            m_TBCLR.Add(m_sUTF2BIG5FilePath, m_oUTF2BIG5TB)
            m_TBStatus.Add(m_sUTF2BIG5FilePath, TBConvertCLR.TBConvertReturnVal.CONVT_SUCCESS)
        Else
            m_TBStatus.Add(m_sUTF2BIG5FilePath, TBConvertCLR.TBConvertReturnVal.CONVT_LOAD_TABLE_FAIL)
        End If
        Return m_TBCLR
    End Function

    Shared Function isLoad(ByVal sPath As String) As Boolean
        '�p�G�ɮת�������O���餤����X�����®� or ���J���Ѯɭ��s���J
        If ST_TB_Header.TBCheckVersion(sPath, m_TBCLR.Item(sPath).m_Header) = 1 OrElse m_TBStatus.Item(sPath) <> TBConvertCLR.TBConvertReturnVal.CONVT_SUCCESS Then
            Return False
        Else
            Return True
        End If
    End Function
 
    '�ˬd�r��O�_�]�t���X�k�r�� 
    Shared Function validateLocalUCS2(ByVal sIn As String, ByVal sOut As String) As Boolean
        Dim pBytes() As Byte = System.Convert.FromBase64String(sIn)

        Dim nRet As TBConvertReturnVal = TBConvertCLR.Validate.TBValidateLocalUCS2(m_oUCS2BIG5TB, sIn, sOut)
        If Convert.ToInt32(nRet) = TBConvertCLR.TBConvertReturnVal.CONVT_INVALID_CHAR Then
            Return False
        End If
        Return True
    End Function

    Shared Function convertUTF8BIG5(ByVal sIn As String) As String
        Dim byOut(65536) As Byte '�j�p���w�q65535�A�ݭn�վ�
        If TBConvertCLR.UTF8Convert.TBConvertUTF8BIG5(m_oUTF2BIG5TB, System.Text.Encoding.Default.GetBytes(sIn), byOut, False) = TBConvertCLR.TBConvertReturnVal.CONVT_SUCCESS Then
            Return System.Text.Encoding.Default.GetString(byOut)
        End If
        Return ""
    End Function
   
End Class
