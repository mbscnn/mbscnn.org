Imports com.Azion.EloanUtility.FileUtility
Public Class ENProperties

    Public Shared ReadOnly m_arrayLoanCase() As String = {"1:新借", "2:續借", "3:增借", "4:減借", "5:改借", "6:新增共用項目", "7:合併續借", "8:合併增借", "9:合併減借", "10:減少項目續借", "11:減少項目增借", "12:減少項目減借", "13:提前到期續借", "14:提前到期增借", "15:提前到期減借", "16:變更", "17:展延", "18:原借"}
    
    ''是否開放擔保品重組
    'Public Shared Function get_ENABLE_GA_REORG() As Boolean
    '    'Return EnConfiguration.get_ENABLE_GA_REORG 
    '    Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("GA_REORG"), Boolean)
    'End Function

    ' ''是否開放更新AP_SCORE_MAIN
    'Public Shared Function get_ENABLE_SCORE() As Boolean
    '    'Return EnConfiguration.get_ENABLE_SCORE 
    '    Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_SCORE"), Boolean)
    'End Function

    ''是否開放下拉專案核定階層
    'Public Shared Function get_ENABLE_LEVEL_OVER_600() As Boolean
    '    'Return EnConfiguration.get_ENABLE_LEVEL_OVER_600 
    '    Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_LEVEL_OVER_600"), Boolean)
    'End Function

    ''是否開放核定後修改共用額度
    'Public Shared Function get_ENABLE_JOIN() As Boolean
    '    'Return EnConfiguration.get_ENABLE_JOIN
    '    Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_JOIN"), Boolean)
    'End Function

    ''是否啟用新版授權檢核條文
    'Public Shared Function get_ENABLE_NEW_LEVELCHK() As String
    '    'Return getAPPValue("ENABLE_NEW_LEVELCHK")
    '    Return CType(com.Azion.EloanUtility.FileUtility.getAppSettings("ENABLE_NEW_LEVELCHK"), Boolean)
    'End Function
End Class


'<!--EN設定 start-->
'<!--<add key="ENProfile" value="C:\eLoanConf\EN_PROFILE.config"/>-->
'  <!--是否開放核定後修改共用額度-->
'  <add key ="ENABLE_JOIN" value="TRUE"/>
'  <!--是否開放下拉專案核定階層-->
'  <add key ="ENABLE_LEVEL_OVER_600" value="TRUE"/>
'  <!--是否開放更新AP_SCORE_MAIN-->
'  <add key ="ENABLE_SCORE" value="TRUE"/>
'  <!--啟用新版授權檢核條文TRUE/FALSE-->
'  <add key ="ENABLE_NEW_LEVELCHK" value="TRUE"/>
'  <!--是否開放更新是否開放擔保品重組-->
'  <add key ="GA_REORG" value="TRUE"/>   
'<!--EN設定End-->
