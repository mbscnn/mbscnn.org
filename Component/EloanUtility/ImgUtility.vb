Imports System.Drawing
Imports System.Drawing.Graphics
Imports System.IO
Imports System.Web

Public Class ImgUtility

    Public Shared ReadOnly UploadFileSize As String = com.Azion.EloanUtility.FileUtility.getAppSettings("UploadFileSize")
    Public Shared ReadOnly ScaleFactor As String = com.Azion.EloanUtility.FileUtility.getAppSettings("ScaleFactor")
    ''' <summary>
    ''' 縮放圖檔
    ''' </summary>
    ''' <param name="scaleFactor">比例</param>
    ''' <param name="sFileName">檔名></param> 
    ''' <remarks>
    ''' [Titan] Create 2011/12/23
    ''' </remarks>
    ''' <returns>MemoryStream</returns>
    Public Shared Function reSizeImage(ByVal scaleFactor As Double, ByVal sFileName As String) As System.IO.MemoryStream 'Bitmap ', ByVal toStream As Stream

        If Not File.Exists(sFileName) Then
            Return Nothing
        End If

        Dim fromStream As New System.IO.MemoryStream(File.ReadAllBytes(sFileName))
        Dim toStream As New System.IO.MemoryStream
        Dim image As System.Drawing.Image = Nothing
        Dim thumbnailBitmap As Bitmap = Nothing
        Dim thumbnailGraph As Graphics = Nothing
        Try
            '----------------------------------------------------------------
            image = System.Drawing.Image.FromStream(fromStream)
            Dim newWidth As Integer = CType((image.Width * scaleFactor), Integer)
            Dim newHeight As Integer = CType((image.Height * scaleFactor), Integer)
            '----------------------------------------------------------------
            thumbnailBitmap = New Bitmap(newWidth, newHeight)
            thumbnailGraph = Graphics.FromImage(thumbnailBitmap)
            '----------------------------------------------------------------
            thumbnailGraph.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            thumbnailGraph.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            thumbnailGraph.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            '----------------------------------------------------------------
            Dim imageRectangle As Rectangle = New Rectangle(0, 0, newWidth, newHeight)
            '----------------------------------------------------------------
            thumbnailGraph.DrawImage(image, imageRectangle)
            thumbnailBitmap.Save(toStream, image.RawFormat)
            '----------------------------------------------------------------
        Catch ex As Exception
            Throw ex
        Finally
            If Not IsNothing(thumbnailGraph) Then thumbnailGraph.Dispose()
            If Not IsNothing(thumbnailGraph) Then thumbnailBitmap.Dispose()
            If Not IsNothing(image) Then image.Dispose()
        End Try

        Return toStream
    End Function


    Shared arraryExtension() As String = {".jpg", ".bmp", ".gif", ".png", ".wmf", ".mix", ".ico"}

    Public Shared Function isImgFile(ByVal sFileName As String) As Boolean
        Dim fileInfo As New System.IO.FileInfo(sFileName)
        If arraryExtension.Contains(fileInfo.Extension) Then Return True
        Return False
    End Function

    Shared Sub checkFileSize(ByVal postedFile As HttpPostedFile)
        If ImgUtility.UploadFileSize <> Nothing Then
            If (postedFile.ContentLength >= CInt(ImgUtility.UploadFileSize)) Then
                Throw New Exception("上傳檔案不能超過" & ImgUtility.UploadFileSize / 1024000 & "M！")
            End If
        End If
    End Sub
End Class
