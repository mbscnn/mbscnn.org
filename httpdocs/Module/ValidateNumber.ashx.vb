Imports System
Imports System.Web
Imports System.Drawing
Imports System.Web.SessionState

Public Class ValidateNumber
    Implements System.Web.IHttpHandler, System.Web.SessionState.IRequiresSessionState

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim NumCount As Integer = 5

        '預設產生5位亂數
        If Not String.IsNullOrEmpty(context.Request.QueryString("NumCount")) Then
            '有指定產生幾位數
            '字串轉數字，轉型成功的話儲存到 NumCount，不成功的話，NumCount會是0
            Int32.TryParse(context.Request.QueryString("NumCount").Replace("'", "''"), NumCount)
        End If

        If NumCount = 0 Then
            NumCount = 5
        End If

        '取得亂數
        Dim str_ValidateCode As String = Me.GetRandomNumberString(NumCount)
        '用於驗證的Session

        context.Session("ValidateNumber") = str_ValidateCode

        '取得圖片物件
        Dim image As System.Drawing.Image = Me.CreateCheckCodeImage(context, str_ValidateCode)
        Dim ms As New System.IO.MemoryStream()
        image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)

        '輸出圖片
        context.Response.Clear()
        context.Response.ContentType = "image/jpeg"
        context.Response.BinaryWrite(ms.ToArray())
        ms.Close()
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

#Region "產生數字亂數"
    Private Function GetRandomNumberString(int_NumberLength As Integer) As String
        Dim str_Number As New System.Text.StringBuilder()

        '字串儲存器
        Dim rand As New Random(Guid.NewGuid().GetHashCode())

        '亂數物件
        Dim i As Integer = 1

        While i <= int_NumberLength
            '產生0~9的亂數
            str_Number.Append(rand.[Next](0, 10).ToString())
            System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While

        Return str_Number.ToString()
    End Function
#End Region

#Region "產生圖片"
    Private Function CreateCheckCodeImage(context As HttpContext, checkCode As String) As System.Drawing.Image

        Dim image As New System.Drawing.Bitmap((checkCode.Length * 20), 40)
        '產生圖片，寬20*位數，高40像素
        Dim g As System.Drawing.Graphics = Graphics.FromImage(image)


        '生成隨機生成器
        Dim random As New Random(Guid.NewGuid().GetHashCode())
        Dim int_Red As Integer = 0
        Dim int_Green As Integer = 0
        Dim int_Blue As Integer = 0
        int_Red = random.[Next](256)
        '產生0~255
        int_Green = random.[Next](256)
        '產生0~255
        int_Blue = (If(int_Red + int_Green > 400, 0, 400 - int_Red - int_Green))
        int_Blue = (If(int_Blue > 255, 255, int_Blue))

        '清空圖片背景色
        g.Clear(Color.FromArgb(int_Red, int_Green, int_Blue))

        '畫圖片的背景噪音線
        Dim i As Integer = 0
        While i <= 24
            Dim x1 As Integer = random.[Next](image.Width)
            Dim x2 As Integer = random.[Next](image.Width)
            Dim y1 As Integer = random.[Next](image.Height)
            Dim y2 As Integer = random.[Next](image.Height)

            g.DrawLine(New Pen(Color.Silver), x1, y1, x2, y2)

            g.DrawEllipse(New Pen(Color.DarkViolet), New System.Drawing.Rectangle(x1, y1, x2, y2))
            System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While

        Dim font As Font = New System.Drawing.Font("Arial", 20, (System.Drawing.FontStyle.Bold))
        Dim brush As New System.Drawing.Drawing2D.LinearGradientBrush(New Rectangle(0, 0, image.Width, image.Height), Color.Blue, Color.DarkRed, 1.2F, True)

        g.DrawString(checkCode, font, brush, 2, 2)
        i = 0
        While i <= 99

            '畫圖片的前景噪音點
            Dim x As Integer = random.[Next](image.Width)
            Dim y As Integer = random.[Next](image.Height)

            image.SetPixel(x, y, Color.FromArgb(random.[Next]()))
            System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While

        '畫圖片的邊框線
        g.DrawRectangle(New Pen(Color.Silver), 0, 0, image.Width - 1, image.Height - 1)

        Return image
    End Function

#End Region

End Class