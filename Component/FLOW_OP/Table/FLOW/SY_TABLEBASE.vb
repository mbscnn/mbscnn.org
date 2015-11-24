Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports FLOW_OP.TABLE

Namespace TABLE


    ''' <summary>
    ''' SY_TABLEBASE are both a base class and factory pattern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SY_TABLEBASE
        Inherits BosBase

        Protected m_SYBranch As SY_BRANCH
        Protected m_SYRelBranchUser As SY_REL_BRANCH_USER
        Protected m_SYUser As SY_USER
        Protected m_SYRole As SY_ROLE
        Protected m_SYRelRoleFlowMap As SY_REL_ROLE_FLOWMAP
        Protected m_SYRelRoleUser As SY_REL_ROLE_USER

        Protected m_SYCaseid As SY_CASEID
        Protected m_SYFlowIncident As SY_FLOWINCIDENT
        Protected m_SYFlowId As FLOW_OP.TABLE.SY_FLOW_ID
        Protected m_SYFlowMap As FLOW_OP.TABLE.SY_FLOW_MAP
        Protected m_SYFlowStep As FLOW_OP.TABLE.SY_FLOWSTEP
        Protected m_SYFlowDef As FLOW_OP.TABLE.SY_FLOW_DEF
        Protected m_SYConditionId As SY_CONDITION_ID
        Protected m_SYCaseResivion As SY_CASEREVISION
        Protected m_SYRelFlowMap_SpanBranch As SY_REL_FLOWMAP_SPANBRANCH

        Protected m_SYNextFLowStepRule As SY_NEXTFLOWSTEPRULE


        Public Sub New(ByVal tableName As String, ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new(tableName, dbManager)
        End Sub

        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.New()
            setDbManager(dbManager)
        End Sub

        Public ReadOnly Property getSYBranch() As SY_BRANCH
            Get
                If IsNothing(m_SYBranch) Then m_SYBranch = New SY_BRANCH(getDatabaseManager())
                Return m_SYBranch
            End Get
        End Property

        ReadOnly Property getSYCaseId() As SY_CASEID
            Get
                If IsNothing(m_SYCaseid) Then m_SYCaseid = New SY_CASEID(getDatabaseManager())
                Return m_SYCaseid
            End Get
        End Property

        ReadOnly Property getSYRelBranchUser() As SY_REL_BRANCH_USER
            Get
                If IsNothing(m_SYRelBranchUser) Then m_SYRelBranchUser = New SY_REL_BRANCH_USER(getDatabaseManager())
                Return m_SYRelBranchUser
            End Get
        End Property

        ReadOnly Property getSYRelRoleUser() As SY_REL_ROLE_USER
            Get
                If IsNothing(m_SYRelRoleUser) Then m_SYRelRoleUser = New SY_REL_ROLE_USER(getDatabaseManager())
                Return m_SYRelRoleUser
            End Get
        End Property

        ReadOnly Property getSYUser() As SY_USER
            Get
                If IsNothing(m_SYUser) Then m_SYUser = New SY_USER(getDatabaseManager())
                Return m_SYUser
            End Get
        End Property

        ReadOnly Property getSYRole() As SY_ROLE
            Get
                If IsNothing(m_SYRole) Then m_SYRole = New SY_ROLE(getDatabaseManager())
                Return m_SYRole
            End Get
        End Property


        ReadOnly Property getSYFlowId() As TABLE.SY_FLOW_ID
            Get
                If IsNothing(m_SYFlowId) Then m_SYFlowId = New FLOW_OP.TABLE.SY_FLOW_ID(getDatabaseManager())
                Return m_SYFlowId
            End Get
        End Property

        ReadOnly Property getSYFlowDef() As TABLE.SY_FLOW_DEF
            Get
                If IsNothing(m_SYFlowDef) Then m_SYFlowDef = New FLOW_OP.TABLE.SY_FLOW_DEF(getDatabaseManager())
                Return m_SYFlowDef
            End Get
        End Property

        ReadOnly Property getSYFlowIncident() As TABLE.SY_FLOWINCIDENT
            Get
                If IsNothing(m_SYFlowIncident) Then m_SYFlowIncident = New FLOW_OP.TABLE.SY_FLOWINCIDENT(getDatabaseManager())
                Return m_SYFlowIncident
            End Get
        End Property

        ReadOnly Property getSYCaseRevision() As TABLE.SY_CASEREVISION
            Get
                If IsNothing(m_SYCaseResivion) Then m_SYCaseResivion = New FLOW_OP.TABLE.SY_CASEREVISION(getDatabaseManager())
                Return m_SYCaseResivion
            End Get
        End Property

        ReadOnly Property getSYFlowMap() As TABLE.SY_FLOW_MAP
            Get
                If IsNothing(m_SYFlowMap) Then m_SYFlowMap = New FLOW_OP.TABLE.SY_FLOW_MAP(getDatabaseManager())
                Return m_SYFlowMap
            End Get
        End Property

        ReadOnly Property getSYFlowStep() As TABLE.SY_FLOWSTEP
            Get
                If IsNothing(m_SYFlowStep) Then m_SYFlowStep = New FLOW_OP.TABLE.SY_FLOWSTEP(getDatabaseManager())
                Return m_SYFlowStep
            End Get
        End Property

        ReadOnly Property getSYConditionId() As TABLE.SY_CONDITION_ID
            Get
                If IsNothing(m_SYConditionId) Then m_SYConditionId = New FLOW_OP.TABLE.SY_CONDITION_ID(getDatabaseManager())
                Return m_SYConditionId
            End Get
        End Property

        ReadOnly Property getSYRelRoleFlowMap() As SY_REL_ROLE_FLOWMAP
            Get
                If IsNothing(m_SYRelRoleFlowMap) Then m_SYRelRoleFlowMap = New SY_REL_ROLE_FLOWMAP(getDatabaseManager())
                Return m_SYRelRoleFlowMap
            End Get
        End Property

        ReadOnly Property getSYRelFlowMap_SpanBranch() As SY_REL_FLOWMAP_SPANBRANCH
            Get
                If IsNothing(m_SYRelFlowMap_SpanBranch) Then
                    m_SYRelFlowMap_SpanBranch = New SY_REL_FLOWMAP_SPANBRANCH(getDatabaseManager())
                End If

                Return m_SYRelFlowMap_SpanBranch
            End Get
        End Property

        ReadOnly Property getSYNextFlowStepRule() As SY_NEXTFLOWSTEPRULE
            Get
                If IsNothing(m_SYNextFLowStepRule) Then
                    m_SYNextFLowStepRule = New SY_NEXTFLOWSTEPRULE(getDatabaseManager())
                End If

                Return m_SYNextFLowStepRule
            End Get
        End Property



        Public Overloads Shared Function CDbType(Of T)(ByVal Expression As Object) As T

            If IsNothing(Expression) OrElse IsDBNull(Expression) Then
                Return Nothing
            End If

            Return CType(Expression, T)
        End Function


        Public Overloads Shared Function CDbType(Of T)(ByVal Expression As Object, ByVal tDefault As T) As T

            If IsNothing(Expression) OrElse IsDBNull(Expression) Then
                Return tDefault
            End If

            Return CType(Expression, T)
        End Function

        Public Shared Function ToArray(Of T)(ByVal ParamArray objs() As T) As T()

            Dim listObjs As New List(Of T)

            For Each obj As T In objs
                listObjs.Add(obj)
            Next

            Return listObjs.ToArray()
        End Function



        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="bosParameter"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function GetDataRow(ByVal bosParameter As BosParameter) As DataRow
            Try
                Return GetDataRow(Nothing,
                    bosParameter.parameterName, bosParameter.parameterValue)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得最後一筆SQL語法
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>在DBUtility更新後可刪除</remarks>
        Public Overloads ReadOnly Property GetLastSQL() As String
            Get
                Return MyBase.GetLastSQL()
            End Get
        End Property

    End Class

End Namespace

