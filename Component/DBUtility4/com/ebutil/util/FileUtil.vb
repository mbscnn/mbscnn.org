Imports System.IO
Imports System.IO.Compression
Imports System.Linq

Public Class FileUtil

#Region "compress"
    Public Shared Sub compressFile(ByVal sourceFile As String, ByVal destinationFile As String)
        '壓縮檔案
        If Not File.Exists(sourceFile) Then
            Throw New FileNotFoundException
        End If

        Dim sourceStream As FileStream = Nothing
        Dim destinationStream As FileStream = Nothing
        Dim compressedStream As GZipStream = Nothing

        Try
            'Read the bytes from the source file into a byte array
            sourceStream = New FileStream(sourceFile, FileMode.Open, FileAccess.Read, FileShare.Read)
            'Read the source stream values into the buffer
            Dim buffer(sourceStream.Length - 1) As Byte
            Dim checkCounter As Integer = sourceStream.Read(buffer, 0, buffer.Length)

            If checkCounter <> buffer.Length Then
                Throw New ApplicationException
            End If

            'Open the FileStream to write to
            destinationStream = New FileStream(destinationFile, FileMode.OpenOrCreate, FileAccess.Write)

            'Create a compression stream pointing to the destiantion stream
            compressedStream = New GZipStream(destinationStream, CompressionMode.Compress, True)

            'Now write the compressed data to the destination file
            compressedStream.Write(buffer, 0, buffer.Length)
        Catch ex As ApplicationException
            Throw ex
        Finally
            'Make sure we allways close all streams
            If sourceStream IsNot Nothing Then
                sourceStream.Close()
            End If

            If compressedStream IsNot Nothing Then
                compressedStream.Close()
            End If

            If destinationStream IsNot Nothing Then
                destinationStream.Close()
            End If
        End Try
    End Sub

    Public Shared Sub decompressFile(ByVal sourceFile As String, ByVal destinationFile As String)
        '解壓縮檔案
        'make sure the source file is there
        If Not File.Exists(sourceFile) Then
            Throw New FileNotFoundException
        End If

        'Create the streams and byte arrays needed
        Dim sourceStream As FileStream = Nothing
        Dim destinationStream As FileStream = Nothing
        Dim decompressedStream As GZipStream = Nothing
        Dim quartetBuffer(4) As Byte

        Try
            'Read in the compressed source stream
            sourceStream = New FileStream(sourceFile, FileMode.Open)

            'Create a compression stream pointing to the destiantion stream
            decompressedStream = New GZipStream(sourceStream, CompressionMode.Decompress, True)

            'Read the footer to determine the length of the destiantion file
            Dim position As Integer = sourceStream.Length - 4
            sourceStream.Position = position
            sourceStream.Read(quartetBuffer, 0, 4)
            sourceStream.Position = 0
            Dim checkLength As Integer = BitConverter.ToInt32(quartetBuffer, 0)

            Dim buffer(checkLength + 100) As Byte
            Dim offset, total As Integer
            'Read the compressed data into the buffer
            While (True)
                Dim bytesRead As Integer = decompressedStream.Read(buffer, offset, 100)

                If bytesRead = 0 Then Exit While

                offset += bytesRead
                total += bytesRead
            End While

            'Now write everything to the destination file
            destinationStream = New FileStream(destinationFile, FileMode.Create)
            destinationStream.Write(buffer, 0, total)

            'and flush everyhting to clean out the buffer
            destinationStream.Flush()

        Catch ex As ApplicationException
            Throw ex
        Finally
            'Make sure we allways close all streams
            If sourceStream IsNot Nothing Then
                sourceStream.Close()
            End If

            If decompressedStream IsNot Nothing Then
                decompressedStream.Close()
            End If

            If destinationStream IsNot Nothing Then
                destinationStream.Close()
            End If
        End Try
    End Sub
#End Region
    
#Region "write file"
    Public Shared Sub writeFile(ByVal sFileName As String, ByVal str As String)
        Dim sw As System.IO.StreamWriter = Nothing
        Try
            sw = System.IO.File.CreateText(sFileName)
            sw.Write(str)
        Catch ex As Exception
            Throw ex
        Finally
            sw.Flush()
            sw.Close()
        End Try
    End Sub

#End Region

#Region "File operatpr"
    Public Shared Function fileList(ByVal sDir As String) As FileSystemInfo()
        Dim di As DirectoryInfo
        Dim files() As FileSystemInfo = Nothing
        Dim orderedFiles As IEnumerable(Of FileSystemInfo) = Nothing
        If My.Computer.FileSystem.DirectoryExists(sDir) Then
            di = New DirectoryInfo(sDir)
            files = di.GetFileSystemInfos()
        End If

        Return files
    End Function

    ''' <summary>
    ''' 取得目錄下副檔名為特定格式的檔案
    ''' </summary>
    ''' <param name="sDirPath">起始目錄</param>
    ''' <param name="sFileExtension">要尋找的副檔名</param>
    ''' <param name="bRecursive">是否要尋找子目錄(預設值為不尋找子目錄)</param>
    ''' <returns>List(Of FileInfo)</returns>
    ''' <remarks></remarks>
    ''' <permission></permission>
    Public Shared Function getPathFilebyExet(ByVal sDirPath As String, _
                                 ByVal sFileExtension As String(), _
                                 Optional ByVal bRecursive As Boolean = False) _
                                 As List(Of FileInfo)

        '取得目錄下所有的資料夾
        Dim DirectoryPath As New List(Of String)
        DirectoryPath = My.Computer.FileSystem.GetDirectories(sDirPath).ToList
        Dim Files As New List(Of FileInfo)

        If bRecursive = True Then
            For Each sDirName As String In DirectoryPath
                If System.IO.Directory.Exists(sDirName) Then

                    '如果存在下一層的資料夾就遞迴呼叫
                    Files.AddRange(getPathFilebyExet(sDirName, sFileExtension, bRecursive))
                End If
            Next
        End If

        DirectoryPath.Add(sDirPath)
        For Each DirStr As String In DirectoryPath
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
            Where sFileExtension.Contains(s.Extension.ToLower)

            Files.AddRange(filter.ToList)
        Next

        Return Files
    End Function


    ''' <summary>
    ''' 取得目錄下相似的檔名檔案
    ''' </summary>
    ''' <param name="sDirPath">起始目錄</param>
    ''' <param name="sNaming">要尋找的副檔名</param>
    ''' <param name="bRecursive">是否要尋找子目錄(預設值為不尋找子目錄)</param>
    ''' <returns>List(Of FileInfo)</returns>
    ''' <remarks></remarks>
    ''' <permission></permission>
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

#Region "Configuration"

    Public Shared Function getAppSettings(ByVal sKey As String, Optional ByVal sFileName As String = Nothing) As String
        Dim sValue As String = System.Configuration.ConfigurationManager.AppSettings(sKey)

        If IsNothing(sValue) And Not IsNothing(sFileName) Then
            Dim map As System.Configuration.ExeConfigurationFileMap = New System.Configuration.ExeConfigurationFileMap()
            map.ExeConfigFilename = sFileName

            Dim configuration As System.Configuration.Configuration = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(map, System.Configuration.ConfigurationUserLevel.None)

            sValue = configuration.AppSettings.Settings.Item(sKey).Value
        End If

        Return sValue
    End Function

#End Region

End Class
