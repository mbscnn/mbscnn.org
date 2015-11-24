Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Data

Public Class FieldValidator
    Private _FiledId As String
    Private _VRules As IList(Of Rules)
    Private _FunName As String
    Private _FormName As String
    Private lFieldValidator As IList(Of FieldValidator) = New List(Of FieldValidator)

    ''' <summary>
    ''' 驗證欄位ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FiledId() As String
        Get
            Return _FiledId
        End Get
        Set(ByVal value As String)
            _FiledId = value
        End Set
    End Property

    ''' <summary>
    ''' 驗證規則
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property VRules() As IList(Of Rules)
        Get
            Return _VRules
        End Get
        Set(value As IList(Of Rules))
            _VRules = value
        End Set
    End Property

    ''' <summary>
    ''' Function 名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FunName() As String
        Get
            Return _FunName
        End Get
        Set(value As String)
            _FunName = value
        End Set
    End Property

    ''' <summary>
    ''' Form 名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FormName() As String
        Get
            Return _FormName
        End Get
        Set(value As String)
            _FormName = value
        End Set
    End Property

    ''' <summary>
    ''' 構造函數
    ''' </summary>
    ''' <param name="sFunName">function 名稱</param>
    ''' <param name="sFormName">form 名稱</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sFunName As String, ByVal sFormName As String)
        Me.FunName = sFunName
        Me.FormName = sFormName
    End Sub

    ''' <summary>
    ''' 構造函數
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' 添加驗證子
    ''' </summary>
    ''' <param name="sFieldID">欄位控件ID</param>
    ''' <param name="rRuleInfo">驗證規則</param>
    ''' <remarks></remarks>
    Sub Add(ByVal sFieldID As String, ByVal ParamArray rRuleInfo As Rules())
        For Each validatorItem As FieldValidator In lFieldValidator
            If validatorItem.FiledId = sFieldID Then
                Throw New Exception(String.Format("{0} 已進行初始化!", sFieldID))
            End If
        Next

        If rRuleInfo Is Nothing Or rRuleInfo.Count = 0 Then
            Throw New Exception(String.Format("{0 沒有驗證規則可以初始化！}", sFieldID))
        End If

        Dim validator As New FieldValidator

        validator.FiledId = sFieldID
        validator.VRules = rRuleInfo

        lFieldValidator.Add(validator)
    End Sub

    ''' <summary>
    ''' 添加驗證子
    ''' </summary>
    ''' <param name="FieldControl">欄位控件ID</param>
    ''' <param name="rRuleInfo">驗證規則</param>
    ''' <remarks></remarks>
    Sub Add(ByVal FieldControl As Control, ByVal ParamArray rRuleInfo As Rules())
        If Not FieldControl Is Nothing Then
            Me.Add(FieldControl.UniqueID, rRuleInfo)
        End If
    End Sub

    ''' <summary>
    ''' 去除驗證器
    ''' </summary>
    ''' <param name="sFieldID"></param>
    ''' <remarks></remarks>
    Sub Remove(ByVal sFieldID As String)
        For Each validator As FieldValidator In lFieldValidator
            If validator.FiledId = sFieldID Then
                lFieldValidator.Remove(validator)
            End If
        Next
    End Sub

    ''' <summary>
    ''' 清空驗證器
    ''' </summary>
    ''' <remarks></remarks>
    Sub Clear()
        lFieldValidator.Clear()
    End Sub

    ''' <summary>
    ''' 生成驗證器(JSON格式)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GenValidator() As String
        If lFieldValidator Is Nothing Or lFieldValidator.Count = 0 Then
            Return String.Format(" function {0}(){{}} ", Me.FunName)
        End If

        Dim sValidatorRule As New StringBuilder
        Dim sValidatorMessage As New StringBuilder

        ' 取出每一個驗證器進行初始化
        For Each validator As FieldValidator In lFieldValidator
            Dim bFlag As Boolean = False

            ' 如果已經初始化其他驗證器，則需加入'，'間隔
            If sValidatorRule.Length <> 0 Then
                sValidatorRule.Append(",")
                sValidatorMessage.Append(",")
            End If

            sValidatorRule.Append(String.Format("{0}:{{", validator.FiledId))
            sValidatorMessage.Append(String.Format("{0}:{{", validator.FiledId))

            ' 初始化每個驗證器的所有的驗證規則
            For Each item In validator.VRules
                If bFlag Then
                    sValidatorRule.Append(",")
                    sValidatorMessage.Append(",")
                Else
                    bFlag = True
                End If

                sValidatorRule.Append(Me.SwitchRule(item))
                sValidatorMessage.Append(String.Format("{0}:""{1}""", item.Rule.ToString(), item.Message))
            Next

            sValidatorRule.Append("}")
            sValidatorMessage.Append("}")
        Next

        Return String.Format(" function {0}(){{ $.removeData($('#{1}')[0], 'validator'); $('#{1}').validate({{rules:{{{2}}},messages:{{{3}}}}});}} ", Me.FunName, Me.FormName, sValidatorRule, sValidatorMessage)
    End Function

    ''' <summary>
    ''' 生成按鈕的註冊驗證并觸發驗證的腳本
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ScriptForButton() As String
        ' 檢核Function Name有沒有初始化，如果沒有初始化丟出異常
        If Me.FunName Is Nothing Then
            Throw New Exception("Function Name 沒有進行初始化!")
        End If
        ' 檢核Form Name有沒有初始化，如果沒有初始化丟出異常
        If Me.FormName Is Nothing Then
            Throw New Exception("Form Name 沒有進行初始化!")
        End If

        ' 返回Button前臺觸發驗證的
        Return String.Format(" {0}();return $('#{1}').valid();", Me.FunName, Me.FormName)
    End Function

    ''' <summary>
    ''' 生成按鈕的註冊驗證并觸發驗證的腳本
    ''' </summary>
    ''' <param name="sMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ScriptForComfirmButton(ByVal sMessage As String) As String
        ' 檢核Function Name有沒有初始化，如果沒有初始化丟出異常
        If Me.FunName Is Nothing Then
            Throw New Exception("Function Name 沒有進行初始化!")
        End If
        ' 檢核Form Name有沒有初始化，如果沒有初始化丟出異常
        If Me.FormName Is Nothing Then
            Throw New Exception("Form Name 沒有進行初始化!")
        End If

        ' 返回Button前臺觸發驗證的
        Return String.Format(" {0}();if($('#{1}').valid()){{return confirm('{2}')}} return false;", Me.FunName, Me.FormName, sMessage)
    End Function

    ''' <summary>
    ''' 構造驗證規則
    ''' </summary>
    ''' <param name="rRule"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SwitchRule(ByVal rRule As Rules) As String
        Dim sResult As String = ""

        Select Case rRule.Rule

            Case ValidateType.UpDate
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.NoUpValue
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.CompareEqules
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.CompareNoEqules
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.EmptyOrLessDay
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.EmptyOrLessMonth
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.EmptyOrUperDay
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.BothBeginEnd
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.CompareDate
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.UperDay
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.IsUse
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.ByteRangeLength
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.MaxLength
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.MinLength
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.InputOne
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.GreatEqual
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.GreatThan
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.LessEqual
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.LessThan
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.IsZero
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.LessThanSpecificDate
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.LessEqualSpecificDate
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.GreatThanSpecificDate
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.GreatEqualSpecificDate
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.JustOneValue
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())
            Case ValidateType.RequiredRadioList
                If rRule.CompareTo Is Nothing Then
                    Throw New Exception(String.Format("驗證規則：{0} 沒有初始化 CompareTo！", rRule.Rule.ToString()))
                End If
                sResult = String.Format("{0} : ""{1}""", rRule.Rule.ToString(), rRule.CompareTo.ToString())

            Case Else
                sResult = String.Format("{0} : true", rRule.Rule.ToString())

        End Select
        Return sResult
    End Function

    ''' <summary>
    ''' 顯示空表頭
    ''' </summary>
    ''' <param name="dgGrid"></param>
    ''' <param name="dtTable"></param>
    ''' <remarks></remarks>
    Shared Sub ShowEmptyGridView(ByVal dgGrid As DataGrid, ByVal dtTable As DataTable)
        If dgGrid Is Nothing Then
            Throw New ArgumentNullException("dataGrid")
        End If

        Dim newRow As DataRow = dtTable.NewRow
        dtTable.Rows.Add(newRow)
        dgGrid.DataSource = dtTable
        dgGrid.DataBind()
        dgGrid.Items(0).Visible = False
    End Sub
End Class
