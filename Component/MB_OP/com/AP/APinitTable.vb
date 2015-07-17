
Imports com.Azion.NET.VB

Public Class APinitTable

    '共用物件,程式執行中請勿修改其值
    '===== AP_xxxx... TABLE ===============================================

    ''政策性貸款 AP_RATEFORPLUSE
    'Friend Shared apRateFoePluses As AP_RATEFORPLUSELIST = Nothing
    'Friend Shared dtAPRateFoePluses As DataTable = Nothing

    ''幣別 AP_CURCODE 
    'Friend Shared apCurcodes As AP_CURCODEList = Nothing
    Friend Shared dtApCurcodes As DataTable = Nothing

    '頁籤 第一層 EN_LINK_TAB 
    'Friend Shared _enLinkTabs As EN_LINK_TABList = Nothing
    Friend Shared dtEnLinkTabs As DataTable = Nothing
    'Friend Shared htSubsys As Hashtable

    '頁籤 第二層 EN_PROG_VER 
    'Friend Shared enProgVers As EN_PROG_VERList = Nothing
    Friend Shared dtEnProgVers As DataTable = Nothing

    '項目別代碼
    'Friend Shared apCodes As AP_CODEList = Nothing
    Friend Shared dtApCodes As DataTable = Nothing

    ''消金政策性貸款 CO_LOANITEMLIST
    'Friend Shared coloanItemLists As CO_LOANITEMLISTList = Nothing
    Friend Shared dtColoanItemList As DataTable = Nothing

    '消金政策性貸款
    'Shared Sub initColoanItemListTable(ByVal dbManager As DatabaseManager)
    '    Try
    '        If IsNothing(APinitTable.dtColoanItemList) Then
    '            Dim bosList As CO_LOANITEMLISTList = New CO_LOANITEMLISTList(dbManager)
    '            bosList.setSQLCondition(" order by loanitemid,parentid,orderno")
    '            bosList.loadAllData()
    '            APinitTable.dtColoanItemList = bosList.getCurrentDataSet.Tables(0)
    '        End If
    '    Catch ex As ProviderException
    '        Throw ex
    '    Catch ex As BosException
    '        Throw ex
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    '頁籤 第一層 EN_LINK_TAB 
    '<System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
    'Shared Sub initEnLinkTabTable(ByVal dbManager As DatabaseManager)

    '    Try
    '        If IsNothing(APinitTable.dtEnLinkTabs) Then
    '            Dim bosList As EN_LINK_TABList = New EN_LINK_TABList(dbManager)
    '            bosList.setSQLCondition(" ORDER BY STEP_NO ,SEQ")
    '            bosList.loadAllData(True)
    '            APinitTable.dtEnLinkTabs = bosList.getCurrentDataSet.Tables(0)
    '        End If
    '    Catch ex As ProviderException
    '        Throw ex
    '    Catch ex As BosException
    '        Throw ex
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    '頁籤 第二層 EN_PROG_VER 
    '<System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
    'Shared Sub initEnProgVerTable(ByVal dbManager As DatabaseManager)
    '    Try
    '        If IsNothing(APinitTable.dtEnProgVers) Then
    '            Dim bosList As EN_PROG_VERList = New EN_PROG_VERList(dbManager)
    '            bosList.loadAllData(True)
    '            APinitTable.dtEnProgVers = bosList.getCurrentDataSet.Tables(0)
    '        End If
    '    Catch ex As ProviderException
    '        Throw ex
    '    Catch ex As BosException
    '        Throw ex
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    '項目別代碼
    <System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
    Shared Sub initApCode(ByVal dbManager As DatabaseManager)
        Try
            If IsNothing(APinitTable.dtApCodes) Then
                Dim bosList As AP_CODEList = New AP_CODEList(dbManager)
                bosList.setSQLCondition("ORDER BY SUBSYSID,SORTNO,UPCODE,CODEID")
                bosList.loadAllData(True)
                APinitTable.dtApCodes = bosList.getCurrentDataSet.Tables(0)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '幣別別代碼
    '<System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
    'Shared Sub initApCurcode(ByVal dbManager As DatabaseManager)
    '    Try
    '        If IsNothing(APinitTable.dtApCurcodes) Then
    '            Dim bosList As AP_CURCODEList = New AP_CURCODEList(dbManager)
    '            bosList.setSQLCondition(" ORDER BY CURNO")
    '            bosList.loadAllData(True)
    '            APinitTable.dtApCurcodes = bosList.getCurrentDataSet.Tables(0)
    '        End If
    '    Catch ex As ProviderException
    '        Throw ex
    '    Catch ex As BosException
    '        Throw ex
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    ''政策性貸款
    '<System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
    'Shared Sub initAPRateFoePluse(ByVal dbManager As DatabaseManager)
    '    Try
    '        If IsNothing(APinitTable.dtAPRateFoePluses) Then
    '            Dim bosList As AP_RATEFORPLUSELIST = New AP_RATEFORPLUSELIST(dbManager)
    '            bosList.loadAllData(True)
    '            APinitTable.dtAPRateFoePluses = bosList.getCurrentDataSet.Tables(0)
    '        End If
    '    Catch ex As ProviderException
    '        Throw ex
    '    Catch ex As BosException
    '        Throw ex
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Shared Sub reset()
        'If Utility.isValidateData(APinitTable.apCodes) Then
        '    APinitTable.apCodes.Reset()
        'End If

        'If Utility.isValidateData(APinitTable.apCurcodes) Then
        '    APinitTable.apCurcodes.Reset()
        'End If

        'If Utility.isValidateData(APinitTable.apRateFoePluses) Then
        '    APinitTable.apRateFoePluses.Reset()
        'End If

        'If Utility.isValidateData(APinitTable.enProgVers) Then
        '    APinitTable.enProgVers.Reset()
        'End If

        'If Utility.isValidateData(APinitTable._enLinkTabs) Then
        '    APinitTable._enLinkTabs.Reset()
        'End If
    End Sub


    Shared Sub clean()
        ''政策性貸款 AP_RATEFORPLUSE
        'APinitTable.apRateFoePluses = Nothing
        'APinitTable.dtAPRateFoePluses = Nothing

        ''幣別 AP_CURCODE 
        'APinitTable.apCurcodes = Nothing
        APinitTable.dtApCurcodes = Nothing

        '頁籤 第一層 EN_LINK_TAB 
        'APinitTable._enLinkTabs = Nothing
        APinitTable.dtEnLinkTabs = Nothing

        '頁籤 第二層 EN_PROG_VER 
        'APinitTable.enProgVers = Nothing
        APinitTable.dtEnProgVers = Nothing

        '項目代碼 Ap_Code
        ' APinitTable.apCodes = Nothing
        APinitTable.dtApCodes = Nothing
    End Sub
End Class
