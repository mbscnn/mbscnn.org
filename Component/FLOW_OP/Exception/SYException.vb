

''' <summary>
''' EXCEPTION包含錯誤碼及額外的物件內容，
''' 可讓取得錯誤的程式依照錯誤碼及額外的物件內容自行決定要顯示什麼錯誤內容
''' </summary>
''' <remarks></remarks>
<Serializable()> Public Class SYException
    Inherits Exception

    Protected m_nMsgCode As SYMSG     '錯誤碼
    Protected m_sLastSQL As String      '最後一個執行的SQL

    Public Property Code() As SYMSG
        Get
            Return m_nMsgCode
        End Get
        Set(ByVal value As SYMSG)
            m_nMsgCode = value
        End Set
    End Property

    Public Property LastSQL() As String
        Get
            Return m_sLastSQL
        End Get
        Set(ByVal value As String)
            m_sLastSQL = value
        End Set
    End Property


    Protected m_Object As Object          '要傳回的物件內容
    Public Property Obj() As Object
        Get
            Return m_Object
        End Get
        Set(ByVal value As Object)
            m_Object = value
        End Set
    End Property


    ''' <summary>
    ''' EXCEPTION，包含要傳回的錯誤碼及額外的物件內容，
    ''' 可讓取得錯誤的程式依照錯誤碼及額外的物件內容自行決定要顯示什麼錯誤內容
    ''' </summary>
    ''' <param name="strMsg">錯誤訊息</param>
    ''' <param name="nMsgCode">錯誤碼</param>
    ''' <param name="sLastSQL">最後一個執行的SQL</param>
    ''' <param name="Obj">額外的物件內容</param>
    ''' <remarks></remarks>    
    Public Sub New(ByVal strMsg As String,
                   Optional ByVal nMsgCode As SYMSG = SYMSG.UNDEFINE,
                   Optional ByVal sLastSQL As String = Nothing,
                   Optional ByVal Obj As Object = Nothing)
        MyBase.New(strMsg)
        m_nMsgCode = nMsgCode
        m_sLastSQL = sLastSQL
        m_Object = Obj
    End Sub


    ''' <summary>
    ''' EXCEPTION，包含要傳回的錯誤碼及額外的物件內容，
    ''' 可讓取得錯誤的程式依照錯誤碼及額外的物件內容自行決定要顯示什麼錯誤內容
    ''' </summary>
    ''' <param name="strMsg">錯誤訊息</param>
    ''' <param name="innerException">innerException</param>
    ''' <param name="nMsgCode">錯誤碼</param>
    ''' <param name="sLastSQL">最後一個執行的SQL</param>
    ''' <param name="Obj">額外的物件內容</param>
    ''' <remarks></remarks>    
    Public Sub New(ByVal strMsg As String,
                   ByVal innerException As Exception,
                   Optional ByVal nMsgCode As SYMSG = SYMSG.UNDEFINE,
                   Optional ByVal sLastSQL As String = Nothing,
                   Optional ByVal Obj As Object = Nothing)
        MyBase.New(strMsg, innerException)
        m_nMsgCode = nMsgCode
        m_sLastSQL = sLastSQL
        m_Object = Obj
    End Sub


    ''' <summary>
    ''' 組物件LIST，可一次傳回多個物件內容
    ''' </summary>
    ''' <param name="objs"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property OBJLIST(ByVal ParamArray objs() As Object) As List(Of Object)
        Get
            Dim listObject As New List(Of Object)

            For i As Integer = 0 To objs.Count - 1
                listObject.Add(objs(i))
            Next i

            Return listObject
        End Get
    End Property
End Class
