Imports System.Drawing
Imports System.Drawing.Graphics
Imports System.IO

Module TestMain
    Public Sub Main(ByVal ParamArray spara() As String)
        'Console.WriteLine(String.Format("{0:$#,##0.00;($#,##0.00);Zero}", 0)) '這個會顯示 Zero 
        'Console.WriteLine(String.Format("{0:$#,##0.00;($#,##0.00);Zero}", 1243.5)) ' 這個會顯示 $1,243.50
        'Console.WriteLine(com.Azion.EloanUtility.NumUtility.str2Amount("Ass,$s09.8-"))
        'Console.WriteLine(com.Azion.EloanUtility.NumUtility.str2Amount("12435678912345.596"))

        'Console.WriteLine(com.Azion.EloanUtility.DateUtility.showTWYYMMDD(Now))

        'Console.WriteLine(com.Azion.EloanUtility.DateUtility.showTWYYMM(Now))
        'Console.WriteLine(com.Azion.EloanUtility.DateUtility.showTWYYMMDD(Now))
        '測試縮圖
        'ResizeImage()
        Dim sMailAddrs As String = String.Empty
        If spara.Length > 0 Then
            sMailAddrs = spara(0)
        End If
        If sMailAddrs = Nothing Then
            sMailAddrs = "S238353@sinyi.com.tw;Mango.lin@sinyi.com.tw;S153559@yahoo.com.tw;Beagle0313@entiebank.com.tw;NeilLin@entiebank.com.tw;Felix@entiebank.com.tw;Leolin@entiebank.com.tw;"
            sMailAddrs = sMailAddrs & "peggy@hondapac.com.tw;"
            sMailAddrs = sMailAddrs & "amyh@azion.com.tw;janehsu@azion.com.tw;michellehsieh36@yahoo.com.tw;titanchen@azion.com.tw"
        End If
        com.Azion.EloanUtility.NetUtility.sendMail( _
                                                                             com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingSubject") _
                                                                            , com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingBody") _
                                                                             , sMailAddrs, com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingFromMail") _
                                                                            , com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingSMTP") _
                                                                             , com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingSMTPAcc") _
                                                                             , com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingSMTPPwd") _
                                                                             , com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingSMTPPort") _
                                                                             , com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingFromName") _
                                                                             , CType(com.Azion.EloanUtility.FileUtility.getAppSettings("OutsourcingSMTPisSSL"), Boolean))
    End Sub


    '測試縮圖
    Sub ResizeImage()
        Dim sFileName As String = "C:\Users\Public\Pictures\Sample Pictures\xxx.jpg"
        Dim fileInfo As New System.IO.FileInfo(sFileName)
        Dim sNewFileName As String = fileInfo.FullName.Replace(fileInfo.Extension, "") & "_thumb" & fileInfo.Extension
        If System.IO.File.Exists(sNewFileName) Then
            System.IO.File.Delete(sNewFileName)
        End If
        Dim toStream As System.IO.MemoryStream = com.Azion.EloanUtility.ImgUtility.ResizeImage(0.1, sFileName)

        Dim thumbnailBitmap As New Bitmap(toStream)
        thumbnailBitmap.Save(sNewFileName)
        thumbnailBitmap.Dispose()

    End Sub
End Module
