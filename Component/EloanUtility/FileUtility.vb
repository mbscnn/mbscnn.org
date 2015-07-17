Option Explicit On
Option Strict On

Imports System.IO

'測試
'"C:\Inetpub\vhosts\mbscnn.org\httpdocs\MBSCConf\MBSC.config"
'正式
'"C:\Inetpub\vhosts\mbscnn.org\httpdocs\MBSCConf\MBSC.config"

'<?xml version="1.0" encoding="utf-8" ?>
'<Configuration>
'    <Application name="Infoimg_IDM_Compare">
'        <InfoImg_DBConnection>Provider=OraOleDB.Oracle;max pool size=350;Password=boteloan;User ID=boteloan;Data Source=117</InfoImg_DBConnection>
'			<!-- InfoImage 資料庫的連線字串-->
'        <IDM_DBConnection>Provider=OraOleDB.Oracle;max pool size=350;Password=boteloan;User ID=boteloan;Data Source=117</IDM_DBConnection>
'			<!-- IDM_WEB 資料庫的連線字串-->
'        <DBConnection>Provider=OraOleDB.Oracle;max pool size=350;Password=boteloan;User ID=boteloan;Data Source=117</DBConnection>
'			<!-- INFOIMAGE_IDM_COMP 的資料庫的連線字串，INFOIMAGE_IDM_COMP是用來儲存比對結果-->

'        <IDM_Web_ELOAN>http://127.0.0.1/IDM_Web_ELOAN/</IDM_Web_ELOAN>
'			<!-- IDM WEB的網址，程式從IDM_WEB下載，用來檢查檔案是存真的存在-->
'        <IDM_Web_OTHERS>http://127.0.0.1/IDM_Web_OTHERS/</IDM_Web_OTHERS>
'			<!-- IDM WEB的網址，程式從IDM_WEB下載，用來檢查檔案是存真的存在-->
'        <LogPath>Log\</LogPath>
'			<!-- 應用程式日誌存放目錄 -->
'        <DisplayLog>True</DisplayLog>
'			<!-- 是否於CONSOLE顯示訊息 -->

'        <LastDate></LastDate>
'        <LastTime></LastTime>
'			<!-- 記錄程式所處理最後一筆相關的日期及時間(如此才知道要從哪一筆繼續移轉)，預設為空白 -->
'			<!--	<LastDate></LastDate>  <LastTime></LastTime>	-->

'        <RunMode>04</RunMode>
'            <!-- #只檢查指定的SubsysID的案件，其它SubsysID的案件不會檢查-->
'            <!-- #可能值為 01, 02, 03, 04, 05, 06 和 07'-->
'    </Application>
'</Configuration>



'''''' Read from XML File
'        Dim _Xml As New XmlHelper("setup.xml")

'        m_InfoImg_DBConnection = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/InfoImg_DBConnection")
'        m_IDM_DBConnection = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/IDM_DBConnection")
'        m_DBConnection = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/DBConnection")

'        m_IDM_Web_ELOAN = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/IDM_Web_ELOAN")
'        m_IDM_Web_OTHERS = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/IDM_Web_OTHERS")

'        m_LogPath = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/LogPath")
'        m_DisplayLog = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/DisplayLog")

'        m_LastDate = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/LastDate")
'        m_LastTime = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/LastTime")

'        m_SubSysid = _Xml.GetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/RunMode")



'''''' Write to XML File
'        Dim _Xml As New XmlHelper("setup.xml")
'        _Xml.SetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/LastDate", m_LastDate)
'        _Xml.SetString("/Configuration/Application[@name='Infoimg_IDM_Compare']/LastTime", m_LastTime)
'        _Xml.Save("setup.xml")
''' <summary>
''' 提供檔案相關函式
''' [Titan] 	2011/07/19	Created
''' </summary>
Public Class FileUtility

#Region "Configuration"
    ''' <summary>
    ''' 取得目前應用程式預設組態的參數
    ''' AppSettingsSection 物件包含組態檔之 appSettings 區段的內容
    ''' com.Azion.EloanUtility.FileUtility.getAppSettings("skey",Application("EloanConf"))
    ''' C:\eLoanConf\eloan_EnTie.config
    ''' </summary>
    ''' <param name="sKey">key</param>
    ''' <param name="sFileName">SString 設定檔的位置</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function getAppSettings(ByVal sKey As String, Optional ByVal sFileName As String = Nothing) As String
        If Not ValidateUtility.isValidateData(sFileName) Then
            Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
            If Not IsNothing(currContext) Then
                sFileName = currContext.Application("EloanConf").ToString
            End If
        End If

        Dim sValue As String = System.Configuration.ConfigurationManager.AppSettings(sKey)

        If IsNothing(sValue) And Not IsNothing(sFileName) Then
            Dim fileMap As System.Configuration.ExeConfigurationFileMap = New System.Configuration.ExeConfigurationFileMap() With {.ExeConfigFilename = sFileName}

            Dim configuration As System.Configuration.Configuration = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, System.Configuration.ConfigurationUserLevel.None)

            If Not IsNothing(configuration.AppSettings.Settings.Item(sKey)) Then
                sValue = configuration.AppSettings.Settings.Item(sKey).Value
            End If
        End If

        Return sValue
    End Function

    ''' <summary>
    ''' 取得目前應用程式預設組態的參數
    ''' AppSettingsSection 物件包含組態檔之 appSettings 區段的內容
    ''' com.Azion.EloanUtility.FileUtility.getAppSettings("skey",Application("EloanConf"))
    ''' </summary>
    ''' <param name="sKey">key</param>
    ''' <param name="sFileName">SString 設定檔的位置</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function setAppSettings(ByVal sKey As String, ByVal sValue As String, Optional ByVal sFileName As String = "C:\Inetpub\vhosts\mbscnn.org\httpdocs\MBSCConf\MBSC.config", Optional ByVal isReplace As Boolean = False) As Boolean
        Dim bResult As Boolean = False
        Dim sOrgValue As String = System.Configuration.ConfigurationManager.AppSettings(sKey)
        Try
            If IsNothing(sOrgValue) And Not IsNothing(sFileName) Then
                Dim fileMap As System.Configuration.ExeConfigurationFileMap = New System.Configuration.ExeConfigurationFileMap() With {.ExeConfigFilename = sFileName}

                Dim configuration As System.Configuration.Configuration = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(fileMap, System.Configuration.ConfigurationUserLevel.None)
                configuration.AppSettings.Settings.Item(sKey).Value = sValue

                If Not IsNothing(configuration.AppSettings.Settings.Item(sKey)) Then
                    If isReplace Then
                        configuration.Save(System.Configuration.ConfigurationSaveMode.Modified, True)
                    Else
                        Dim fileinfo As FileInfo = New FileInfo(sFileName)
                        If Not System.IO.Directory.Exists(fileinfo.DirectoryName & "\new") Then System.IO.Directory.CreateDirectory(fileinfo.DirectoryName & "\new")

                        Dim sNewFileName As String = fileinfo.DirectoryName & "\new\" & fileinfo.Name

                        If Not System.IO.File.Exists(sNewFileName) Then
                            configuration.SaveAs(sNewFileName, System.Configuration.ConfigurationSaveMode.Modified, False)
                        Else
                            configuration = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(New System.Configuration.ExeConfigurationFileMap() With {.ExeConfigFilename = sNewFileName}, System.Configuration.ConfigurationUserLevel.None)
                            configuration.AppSettings.Settings.Item(sKey).Value = sValue
                            configuration.Save(System.Configuration.ConfigurationSaveMode.Modified, True)
                        End If
                    End If
                    bResult = True
                End If
            Else
                System.Configuration.ConfigurationManager.AppSettings(sKey) = sValue
                bResult = True
            End If
        Catch ex As Exception
        End Try

        Return bResult
    End Function
#End Region

#Region "File operatpr"
    ''' <summary>
    ''' 取得目錄下副檔名為特定格式的檔案
    ''' </summary>
    ''' <param name="sDirPath">起始目錄</param>
    ''' <param name="sFileExtension">要尋找的副檔名</param>
    ''' <param name="bRecursive">是否要尋找子目錄(預設值為不尋找子目錄)</param>
    ''' <returns>List(Of FileInfo)</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function getPathFilebyExet(ByVal sDirPath As String, _
                                 ByVal sFileExtension As String(), _
                                 Optional ByVal bRecursive As Boolean = False) _
                                 As List(Of FileInfo)

        '取得目錄下所有的資料夾
        Dim listDirectoryPath As New List(Of String)
        Dim listFile As New List(Of FileInfo)

        If bRecursive = True Then
            listDirectoryPath = My.Computer.FileSystem.GetDirectories(sDirPath).ToList
            For Each sDirName As String In listDirectoryPath
                If System.IO.Directory.Exists(sDirName) Then

                    '如果存在下一層的資料夾就遞迴呼叫
                    listFile.AddRange(getPathFilebyExet(sDirName, sFileExtension, bRecursive))
                End If
            Next
        End If

        listDirectoryPath.Add(sDirPath)
        For Each sDir As String In listDirectoryPath
            '取得目錄下所有的檔案名稱(String)
            Dim myFiles = From s In My.Computer.FileSystem.GetFiles(sDir)
            '先把檔案名稱轉成FileInfo
            Dim f As New List(Of FileInfo)
            For Each s As String In myFiles
                f.Add(My.Computer.FileSystem.GetFileInfo(s))
            Next

            '使用LINQ來取出我們要的資料(副檔名包含在ImageExtension()裡面的)
            Dim filter As IEnumerable(Of FileInfo) = _
            From s In f _
            Where sFileExtension.Contains(s.Extension.ToLower)

            listFile.AddRange(filter.ToList)
        Next

        Return listFile
    End Function


    ''' <summary>
    ''' 取得目錄下相似的檔名檔案
    ''' </summary>
    ''' <param name="sDirPath">起始目錄</param>
    ''' <param name="sNaming">要尋找的檔名</param>
    ''' <param name="bRecursive">是否要尋找子目錄(預設值為不尋找子目錄)</param>
    ''' <returns>List(Of FileInfo)</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function getPathFilebyNaming(ByVal sDirPath As String, _
                                 ByVal sNaming As String, _
                                 Optional ByVal bRecursive As Boolean = False) _
                                 As List(Of FileInfo)

        '取得目錄下所有的資料夾
        Dim lists As New List(Of String)
        lists = My.Computer.FileSystem.GetDirectories(sDirPath).ToList
        Dim Files As New List(Of FileInfo)

        If bRecursive = True Then
            For Each sDirName As String In lists
                If System.IO.Directory.Exists(sDirName) Then

                    '如果存在下一層的資料夾就遞迴呼叫
                    Files.AddRange(getPathFilebyNaming(sDirName, sNaming, bRecursive))
                End If
            Next
        End If

        lists.Add(sDirPath)
        For Each DirStr As String In lists
            '取得目錄下所有的檔案名稱(String)
            Dim myFiles = From s In My.Computer.FileSystem.GetFiles(DirStr)
            '先把檔案名稱轉成FileInfo
            Dim f As New List(Of FileInfo)
            For Each s As String In myFiles
                f.Add(My.Computer.FileSystem.GetFileInfo(s))
            Next

            '使用LINQ來取出我們要的資料(副檔名包含在ImageExtension()裡面的)
            Dim filter As IEnumerable(Of FileInfo) = _
            From s In f _
            Where s.Name.IndexOf(sNaming) <> -1

            Files.AddRange(filter.ToList)
        Next

        Return Files
    End Function
#End Region
End Class
