Imports FLOW_OP
Imports com.Azion.EloanUtility

Public Class ExportUserInfo
    Implements IFlowCallBack

    Public Overridable Sub GetCurrentUserid(ByRef loginUserid As String, ByRef workingUserid As String) Implements IFlowCallBack.GetCurrentUserid

        Dim g_oUserInfo As New StaffInfo
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current

        g_oUserInfo = CType(currContext.Session("StaffInfo"), StaffInfo)
        workingUserid = g_oUserInfo.WorkingStaffid
        loginUserid = g_oUserInfo.LoginUserId

    End Sub

End Class
