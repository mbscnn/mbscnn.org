Public Class AuthProperties
    '啟用 人員送件檢核
    Public Shared Function ENABLE_STAFFID_SEND_VALID() As Boolean
        Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_STAFFID_SEND_VALID"), Boolean)
    End Function

    '啟用 人員分案檢核
    Public Shared Function ENABLE_STAFFID_DIV_VALID() As Boolean
        Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_STAFFID_DIV_VALID"), Boolean)
    End Function

    '取得 遠距
    Public Shared Function getREMOTE_FLAG() As String
        Return com.Azion.EloanUtility.FileUtility.getAppSettings("REMOTE_FLAG")
    End Function

    '是否開放 default.aspx 安控檢核
    Public Shared Function ENABLE_DEFAULT_ASPX_SECURITY() As Boolean
        Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_DEFAULT_ASPX_SECURITY"), Boolean)
    End Function

    'Default.aspx login password
    Public Shared Function get_DEFAULT_ASPX_PWD() As String
        Return com.Azion.EloanUtility.FileUtility.getAppSettings("DEFAULT_ASPX_PWD")
    End Function

    '是否啟用 遠距
    Public Shared Function enableREMOTE() As Boolean
        Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_REMOTE"), Boolean)
    End Function
     
    '# 跑馬燈 速度 10-400 #
    Public Shared Function get_MARQUEE_SCROLLDELAY() As String
        Return com.Azion.EloanUtility.FileUtility.getAppSettings("MARQUEE_SCROLLDELAY")
    End Function

End Class
