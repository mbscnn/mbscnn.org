<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ErrorPage.aspx.vb" Inherits="MBSC.ErrorPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Error Message</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" --> <!--CSS-->
</head>
<body>
    <form id="form1" runat="server">
    <div>
            <!-- #include virtual="~/inc/eLoanTableStart.inc" --> <!--顯示上方外框線-->
			<!-- #include virtual="~/inc/MBSCErrorMsg.inc" --> <!--錯誤訊息區-->
            <!--#include virtual="~/inc/eLoanTableEnd.inc"--> <!--顯示下方外框線-->
    </div>
    </form>
</body>
</html>
