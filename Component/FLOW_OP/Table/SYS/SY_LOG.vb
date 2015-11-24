Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports FLOW_OP.TABLE


Public Enum LOGTYPE As Integer
    TYPE_ERROR = &H0
    TYPE_INFORMATION = &H1
    TYPE_WARNING = &H2
End Enum



Public Class SY_LOG
    Inherits SY_TABLEBASE
    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_LOG", dbManager)
    End Sub



    Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As SY_LOG
        Return New SY_LOG(dbManager)
    End Function


    Public Sub WriteLog(ByVal sMessage As String,
                          Optional ByVal sSource As String = Nothing,
                          Optional ByVal logType As LOGTYPE = LOGTYPE.TYPE_INFORMATION)
        Try
            Insert("LOGTIME", Now,
                   "SOURCE", sSource,
                   "MESSAGE", sMessage,
                   "LOGTYPE", logType)
        Catch ex As Exception
        End Try
    End Sub


    Public Shared Sub WriteLog(ByVal dbManager As DatabaseManager,
                                 ByVal sMessage As String,
                                 Optional ByVal sSource As String = Nothing,
                                 Optional ByVal logType As LOGTYPE = LOGTYPE.TYPE_INFORMATION)
        Try

            If String.IsNullOrEmpty(sMessage) = False Then
                sMessage = Left(sMessage, 8000)
            End If

            getNewInstance(dbManager).Insert("LOGTIME", Now,
                                             "SOURCE", sSource,
                                             "MESSAGE", sMessage,
                                             "LOGTYPE", logType)
        Catch ex As Exception
        End Try
    End Sub



End Class


