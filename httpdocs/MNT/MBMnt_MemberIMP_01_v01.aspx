<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_MemberIMP_01_v01.aspx.vb" Inherits="MBSC.MBMnt_MemberIMP_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
    <title>學員資料匯入</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!--CSS-->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <!--JS-->
</head>
<body>
    <form id="form1" runat="server">
    <!-- #include virtual="~/inc/PageTab.inc" -->
<table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
    <tr>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:185px" >
            <!-- #include virtual="~/inc/vTab.inc" -->
        </td>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
            <!-- #include virtual="~/inc/Signin.inc" -->
            <div align="center">
                <h2 style="color: #800080; font-family: 標楷體">
                    學員資料匯入
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <asp:FileUpload ID="FUEXCEL" runat="server" />
            <asp:Button ID="btIMPORT" runat="server" CssClass="bt" Text="資料匯入" />
            &nbsp;
            <asp:Button ID="btLoadIMP" runat="server" CssClass="bt" Text="載入已匯入資料" />
            <asp:PlaceHolder ID="PLH_RESULT" runat="server" Visible=false>
                <asp:PlaceHolder ID="PLH_Page" runat="server">
                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                        <tr>
                            <td class="mtr" nowrap>
                                目前在第
                                <asp:Label ID="lblCurrentPage" runat="server" Style="color: red" />
                                頁&nbsp;|&nbsp;共
                            </td>
                            <td class="mtr" nowrap>
                                <asp:Label ID="lblTotalPage" runat="server" Style="color: red" />
                                頁
                            </td>
                            <td class="mtr" nowrap>
                                (每頁
                                <asp:Label ID="lblPageSize" runat="server" Style="color: red" />
                                筆) 共
                                <asp:Label ID="lblTotleSize" runat="server" Style="color: red" />
                                筆
                            </td>
                            <td class="mtr" nowrap align="right" width="100%">
                                <asp:ImageButton ID="btnImgFirst" runat="server" Height="16" Width="16" BorderWidth="0"
                                    ImageUrl="~/img/first2.gif" />
                            </td>
                            <td class="mtr" nowrap>
                                &nbsp;|&nbsp;
                                <asp:LinkButton ID="btnLinkFirst" runat="server" Text="第一頁" />&nbsp;|&nbsp;
                            </td>
                            <td class="mtr" nowrap>
                                <asp:ImageButton ID="btnImgPrev" runat="server" Height="16" Width="16" BorderWidth="0"
                                    ImageUrl="~/img/prev2.gif" />
                            </td>
                            <td class="mtr" nowrap>
                                &nbsp;|&nbsp;
                                <asp:LinkButton ID="btnLinkPrev" runat="server" Text="上一頁" />&nbsp;|&nbsp;
                            </td>
                            <td class="mtr" nowrap>
                                <asp:LinkButton ID="btnLinkNext" runat="server" Text="下一頁" />&nbsp;|&nbsp;
                            </td>
                            <td class="mtr" nowrap>
                                <asp:ImageButton ID="btnImgNext" runat="server" Height="16" Width="16" BorderWidth="0"
                                    ImageUrl="~/img/next2.gif" />
                            </td>
                            <td class="mtr" nowrap>
                                &nbsp;|&nbsp;
                                <asp:LinkButton ID="btnLinkLast" runat="server" Text="最後一頁" />&nbsp;|&nbsp;
                            </td>
                            <td class="mtr" nowrap>
                                <asp:ImageButton ID="btnImgLast" runat="server" Height="16" Width="16" BorderWidth="0"
                                    ImageUrl="~/img/last2.gif" />
                            </td>
                            <td class="mtr" nowrap>
                                &nbsp;| 第<asp:DropDownList ID="DDL_Page" runat="server" AutoPostBack="true" />
                                頁
                            </td>
                        </tr>
                    </table>
                </asp:PlaceHolder>
                <table class="CRTable_Top" cellspacing="0" width="100%">
                    <tr>
                        <td width="8%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="10%" class="th1c_b">
                            姓名
                        </td>
                        <td width="10%" class="th1c_b">
                            出生年月日(YYYYMMDD)
                        </td>
                        <td width="10%" class="th1c_b">
                            身分證字號
                        </td>
                        <td width="10%" class="th1c_b">
                            手機
                        </td>
                        <td width="10%" class="th1c_b">
                            電話
                        </td>
                        <td width="10%" class="th1c_b">
                            e-mail
                        </td>
                        <td width="15%" class="th1c_b">
                            通訊地址
                        </td>
                        <td width="10%" class="th1c_b">
                            修改
                        </td>
                        <td width="7%" class="th1c_b">
                            刪除
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_RESULT" runat="server" >
                        <ItemTemplate>
                            <tr>
                                <td width="8%" class="td2c_b">
                                    <!--會員編號-->
                                    <asp:Literal ID="LTL_MB_MEMSEQ" runat="server" Text='<%#getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"),Container.DataItem("MB_AREA"))%>' />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--姓名-->
                                    <asp:Literal ID="LTL_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>' />
                                    <asp:TextBox ID="TXT_MB_NAME" runat="server" Visible=false CssClass="bordertxt" Text='<%#Container.DataItem("MB_NAME")%>' />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--出生年月日-->
                                    <asp:Literal ID="LTL_MB_BIRTH" runat="server" Text='<%#getROCDate(Container.DataItem("MB_BIRTH"))%>' />
                                    <asp:TextBox ID="TXT_MB_BIRTH" runat="server" Visible=false CssClass="bordernum" Text='<%#getROCDate(Container.DataItem("MB_BIRTH"))%>' />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--身分證字號-->
                                    <asp:Literal ID="LTL_MB_ID" runat="server" Text='<%#Container.DataItem("MB_ID")%>' />
                                    <asp:TextBox ID="TXT_MB_ID" runat="server" Visible=false CssClass="bordertxt" Text='<%#Container.DataItem("MB_ID")%>' />
                                </td>
                                <td width="10%" class="td2_b">
                                    <!--手機-->
                                    <asp:Literal ID="LTL_MB_MOBIL" runat="server" Text='<%#Container.DataItem("MB_MOBIL")%>' />
                                    <asp:TextBox ID="TXT_MB_MOBIL" runat="server" Visible=false CssClass="bordertxt" Text='<%#Container.DataItem("MB_MOBIL")%>' />
                                </td>
                                <td width="10%" class="td2_b">
                                    <!--電話-->
                                    <asp:Literal ID="LTL_MB_TEL" runat="server" Text='<%#Container.DataItem("MB_TEL")%>' />
                                    <asp:TextBox ID="TXT_MB_TEL" runat="server" Visible=false CssClass="bordertxt" Text='<%#Container.DataItem("MB_TEL")%>' />
                                </td>
                                <td width="10%" class="td2_b">
                                    <!--e-mail-->
                                    <asp:Literal ID="LTL_MB_EMAIL" runat="server" Text='<%#Container.DataItem("MB_EMAIL")%>' />
                                    <asp:TextBox ID="TXT_MB_EMAIL" runat="server" Visible=false CssClass="bordertxt" Text='<%#Container.DataItem("MB_EMAIL")%>' />
                                </td>
                                <td width="15%" class="td2_b">
                                    <!--通訊地址-->
                                    <asp:Literal ID="LTL_ARRD" runat="server" />
                                    <asp:DropDownList ID="DDL_MB_CITY" runat=server Visible=false DataTextField="CITY" DataValueField="CITY_ID" AutoPostBack=true OnSelectedIndexChanged="DDL_MB_CITY_OnSelectedIndexChanged" />
                                    <asp:DropDownList ID="DDL_MB_VLG" runat=server Visible=false DataTextField="AREA" DataValueField="AREA_ID" />
                                    <asp:TextBox ID="TXT_MB_ADDR" runat="server" Visible=false CssClass="bordertxt" Text='<%#Container.DataItem("MB_ADDR")%>' Width="99%" />
                                    <input type="hidden" id="HID_MB_CITY" runat="server" value='<%#Container.DataItem("MB_CITY")%>' />
                                    <input type="hidden" id="HID_MB_VLG" runat="server" value='<%#Container.DataItem("MB_VLG")%>' />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--修改-->
                                    <asp:ImageButton ID="IMG_EDIT" runat="server" CommandName="EDIT" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                    <asp:Button ID="btYES" runat="server" Text="確定" CssClass="bt" Visible=false CommandName="YES" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                    <asp:Button ID="btCancel" runat="server" Text="取消" CssClass="bt" Visible=false CommandName="CANCEL" />
                                </td>
                                <td width="7%" class="td2c_b">
                                    <asp:ImageButton ID="IMG_DEL" runat="server" CommandName="DELETE" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                </td>
                            </tr>                    
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </asp:PlaceHolder>
        </td>
    </tr>
</table>
    </form>
</body>
</html>
