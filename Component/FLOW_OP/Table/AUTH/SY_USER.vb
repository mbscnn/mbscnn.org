Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text


Public Class USER_ID_NAME
    Public userId As String
    Public userName As String

    Public Shared Function USER(ByVal userId As String, ByVal userName As String) As USER_ID_NAME
        Dim userIdName As New USER_ID_NAME
        userIdName.userId = userId
        userIdName.userName = userName
        Return userIdName
    End Function
End Class


' ''' <summary>
' ''' 使用者狀態 0在職 9離職
' ''' </summary>
' ''' <remarks></remarks>
'Public Enum ENUM_USER_STATUS
'    ONJOB = 0
'    LEFT = 9
'    ERROR_EXCEPTION = -1
'    ERROR_UNKNOWNSTATUS = -2
'End Enum

Namespace TABLE

    ''' <summary>
    ''' 使用者物件
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SY_USER
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_USER", dbManager)
        End Sub


        ' ''' <summary>
        ' ''' 取得使用者的狀態 (使用者狀態 0在職 9離職)
        ' ''' </summary>
        ' ''' <param name="strStaffid">若不設，則為Constructor的Staffid</param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Function GetStatus(Optional ByVal strStaffid As String = Nothing) As ENUM_USER_STATUS

        '    If IsNothing(strStaffid) Then
        '        strStaffid = m_strStaffid
        '    End If

        '    Dim strResult As String
        '    strResult = CStr(ExecuteScalar( _
        '        "SELECT STATUS FROM SY_USER WHERE STAFFID = @STAFFID@ ", _
        '        New CmdParameter() {PARAMETER("STAFFID", strStaffid)}))

        '    Select Case strResult
        '        Case "0"
        '            Return ENUM_USER_STATUS.ONJOB
        '        Case "9"
        '            Return ENUM_USER_STATUS.LEFT
        '    End Select

        '    Return ENUM_USER_STATUS.ERROR_UNKNOWNSTATUS
        'End Function


        ''' <summary>
        ''' 取得使用者的所有資料
        ''' </summary>
        ''' <param name="strStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserInfo(ByVal strStaffid As String) As DataRow

            Dim dr As DataRow
            Try
                dr = GetDataRow(Nothing, "STAFFID", strStaffid)

                If IsNothing(dr) Then
                    Throw New SYException(
                        String.Format("無法從取得SY_USER的內容，STAFFID={0}", strStaffid),
                        SYMSG.SYUSER_USER_NOT_FOUND,
                        GetLastSQL)
                End If

                Return dr

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 取得使用者的名稱及編號
        ''' </summary>
        ''' <param name="strStaffid">DB上的任一欄位</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserIdName(ByVal strStaffid As String) As USER_ID_NAME

            Dim dr As DataRow
            Dim userIdName As USER_ID_NAME = Nothing

            If String.IsNullOrEmpty(strStaffid) Then
                Return Nothing
            End If

            Try
                dr = GetUserInfo(strStaffid)

                userIdName = New USER_ID_NAME
                userIdName.userName = CStr(dr("USERNAME"))
                userIdName.userId = CStr(dr("STAFFID"))

                Return userIdName

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 使用者是否在職
        ''' </summary>
        ''' <param name="strStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsValidUser(ByVal strStaffid As String) As Boolean
            Dim nValue As Integer
            Try
                nValue = GetCount(BosBase.PARAM_ARRAY("STAFFID", strStaffid, "STATUS", 0))

                If nValue > 0 Then
                    Return True
                Else
                    Return False
                End If

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 取得EMail帳號
        ''' </summary>
        ''' <param name="sStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetEMail(ByVal sStaffid As String) As String
            Dim dr As DataRow
            Try
                dr = GetDataRow(Nothing, "STAFFID", sStaffid)

                If IsNothing(dr) Then
                    Return ""
                End If

                Return CDbType(Of String)(dr("EMAIL"), "")

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得上層管理角色的使用者
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetManger(ByVal sStaffid As String, Optional ByVal sFilter As String = Nothing) As DataTable

            Dim sb As StringBuilder

            Try
                sb = New StringBuilder( _
                        "with RC(ROLEID, " & vbCrLf & _
                        "ROLEMGR, " & vbCrLf & _
                        "BRA_DEPNO, " & vbCrLf & _
                        "STAFFID, " & vbCrLf & _
                        "LEVEL) as " & vbCrLf & _
                        " (select SY_ROLE.ROLEID, ROLEMGR, BRA_DEPNO, STAFFID, 0 as LEVEL " & vbCrLf & _
                        "    from SY_ROLE " & vbCrLf & _
                        "   inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                        "      on RU.ROLEID = SY_ROLE.ROLEID " & vbCrLf & _
                        "   where SY_ROLE.DISABLED = '0' " & vbCrLf & _
                        "     and STAFFID = @STAFFID@ " & vbCrLf & _
                        "  union all " & vbCrLf & _
                        "  select A.ROLEID, " & vbCrLf & _
                        "         A.ROLEMGR, " & vbCrLf & _
                        "         RU.BRA_DEPNO, " & vbCrLf & _
                        "         RU.STAFFID, " & vbCrLf & _
                        "         RC.LEVEL + 1 as LEVEL " & vbCrLf & _
                        "    from SY_ROLE as A " & vbCrLf & _
                        "   inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                        "      on RU.ROLEID = A.ROLEID " & vbCrLf & _
                        "   inner join RC " & vbCrLf & _
                        "      on RC.ROLEMGR = A.ROLEID " & vbCrLf & _
                        "   where A.DISABLED = '0') " & vbCrLf & _
                        "select distinct LEVEL, RO.ROLEID, RO.ROLENAME, BR.BRA_DEPNO, BR.BRID, BR.BRCNAME, UR.STAFFID, UR.USERNAME " & vbCrLf & _
                        "  from RC " & vbCrLf & _
                        " inner join SY_ROLE RO " & vbCrLf & _
                        "    on RO.ROLEID = RC.ROLEID " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RC.BRA_DEPNO " & vbCrLf & _
                        " inner join SY_USER UR " & vbCrLf & _
                        "    on UR.STAFFID = RC.STAFFID " & vbCrLf)

                If String.IsNullOrEmpty(sFilter) = False Then
                    sb.Append(" where " & sFilter & vbCrLf)
                End If

                Return GetDataTable(sb.ToString, "STAFFID", sStaffid)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

    End Class

End Namespace
