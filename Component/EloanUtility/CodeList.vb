Option Explicit On
Option Strict On

Public Class CodeList
    Private Shared m_sConfig As String = FileUtility.getAppSettings("CodeListConfigPath")

    Public Shared Function getAppSettings(ByVal sKey As String) As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        If IsNothing(currContext) Then 
            m_sConfig = "C:\MBSCConf\CodeList_MBSC.config"
        End If
        Return FileUtility.getAppSettings(sKey, m_sConfig)
    End Function
End Class
