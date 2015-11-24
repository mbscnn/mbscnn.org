Imports System.Web.SessionState

Public Class Global_asax
    Inherits System.Web.HttpApplication

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 在應用程式啟動時引發
        '正式環境(雲端)
        Dim sConfigPath As String = "C:\Inetpub\vhosts\mbscnn.org\httpdocs\MBSCConf\MBSC.config"
        '測試環境(台科大)
        'Dim sConfigPath As String = "D:\MBSC\MBSCConf\MBSC.config"

        com.Azion.NET.VB.Properties.setConfiguration(sConfigPath)

        'allow zero datetime=true
        '解决 unable to convert MySQL date/time value to System.DateTime
        '.NET DateTime：2007/01/01 12:00:00
        'MySQL DateTime：2007-01-01 12:00:00
        '只要將年月日中間的「/」符號換成「-」即可囉!
        Application("BotDSN") = "allow zero datetime=true;server=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBSource", sConfigPath) &
                                ";database=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DataBase", sConfigPath) &
                                ";uid=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBUserID", sConfigPath) &
                                ";pwd=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBPassword", sConfigPath) &
                                ";CharSet=utf8;"

        Application("EmpDSN") = "allow zero datetime=true;server=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBSource", sConfigPath) &
                                ";database=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DataBase", sConfigPath) &
                                ";uid=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBUserID", sConfigPath) &
                                ";pwd=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBPassword", sConfigPath) &
                                ";CharSet=utf8;"

        '正式環境(雲端)
        Application("EloanConf") = "C:\Inetpub\vhosts\mbscnn.org\httpdocs\MBSCConf\MBSC.config"
        '測試環境(台科大)
        'Application("EloanConf") = "D:\MBSC\MBSCConf\MBSC.config"
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 在工作階段啟動時引發
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 在每個要求開頭引發
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 在一開始嘗試驗證使用時引發
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' 在錯誤發生時引發
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 在工作階段結束時引發
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 在應用程式結束時引發
    End Sub

End Class