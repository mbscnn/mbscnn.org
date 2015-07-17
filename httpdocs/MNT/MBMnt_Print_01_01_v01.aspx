<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Print_01_01_v01.aspx.vb" Inherits="MBSC.MBMnt_Print_01_01_v01_aspx" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html xmlns:o = "urn:schemas-microsoft-com:office:office" xmlns:w = "urn:schemas-microsoft-com:office:word">
<head >
    <title>收據列印</title>
    <!-- #include virtual="~/Css/ENPrintXML.inc" -->
</head>
<body  TopMargin="0" lang=ZH-TW style='tab-interval:24.0pt;text-justify-trim:punctuation' >
    <div class=Section1_Page style='LAYOUT-GRID:18pt none;Vertical-Align:Top;' Runat="Server" id="DIVMAIN">
        <table border=1 cellspacing=0 cellpadding=0 width="100%" style="border-collapse:collapse;mso-padding-alt:0cm 0cm 0cm 0cm" >
            <tr>
                <td width="15%" >
                    法名/姓名
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_NAME" runat="server" />
                </td>
                <td width="15%" >
                    會員編號
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_MEMSEQ" runat="server" />
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    通訊地址
                </td>
                <td width="85%" colspan=3 >
                    <asp:Literal ID="MB_ADDR" runat="server" />
                </td>                
            </tr>
            <tr>
                <td width="15%" >
                    繳款日期
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_TX_DATE" runat="server" />
                </td>
                <td width="15%" >
                    功德項目
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_ITEMID" runat="server" />
                    <asp:PlaceHolder ID="PLH_MB_MEMTYP" runat="server" Visible=false >
                    &nbsp;
                    會員類別：
                    <asp:Literal ID="MB_MEMTYP" runat="server" />
                    </asp:PlaceHolder>
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    繳款方式
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_FEETYPE" runat="server" />
                </td>
                <td width="15%" >
                    繳款金額
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_TOTFEE" runat="server" />
                    <asp:PlaceHolder ID="PLH_MB_TOTFEE_MM" runat="server" Visible=false>
                        每月金額：
                        <asp:Literal ID="MB_TOTFEE_MM" runat="server" />
                    </asp:PlaceHolder>
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    收據捐款名稱
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_RECNAME" runat="server" />
                </td>
                <td width="15%" >
                    所屬區
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_AREA" runat="server" />
                    &nbsp;
                    委員：
                    <asp:Literal ID="MB_LEADER" runat="server" />
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    會費期間
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_MEMFEE" runat="server" />
                    <asp:Literal ID="MB_DESC" runat="server" Visible=false />
                </td>
                <td width="15%" >
                    付款方式
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_PAY_TYPE" runat="server" />
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    票據到期日
                </td>
                <td width="35%" >
                    <asp:Literal ID="NOTE_DUE_DATE" runat="server" />
                </td>
                <td width="15%" >
                    票據號碼
                </td>
                <td width="35%" >
                    <asp:Literal ID="NOTE_NO" runat="server" />
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    發票行
                </td>
                <td width="35%" >
                    <asp:Literal ID="NOTE_BANK" runat="server" />
                    銀行
                    <asp:Literal ID="NOTE_BR" runat="server" />
                    分行
                </td>
                <td width="15%" >
                    發票人
                </td>
                <td width="35%" >
                    <asp:Literal ID="NOTE_HOLDER" runat="server" />
                </td>
            </tr>
            <tr>
                <td width="15%" >
                    票據金額
                </td>
                <td width="35%" >
                    <asp:Literal ID="NOTE_AMT" runat="server" />
                </td>
                <td width="15%" >
                    補印次數
                </td>
                <td width="35%" >
                    <asp:Literal ID="MB_REISU" runat="server" />
                </td>
            </tr>
        </table>                  
    </div>
</body>
</html>
