<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_APV_01_v01.aspx.vb" Inherits="MBSC.MBMnt_APV_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>MBSC核准入會申請單作業</title>
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
                            核准/不核准入會作業
                        </h2>
                    </div>
                    <!--顯示上方外框線-->
                    <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
                    <!--錯誤訊息區-->
                    <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
                    <asp:PlaceHolder ID="PLH_OPTYPE" runat="server" >
                    <table class="CRTable_Top" width="100%" Cellspacing="0">
                        <tr>
                            <td width="10%" class="th1c_b">
                                作業模式
                            </td>
                            <td width="90%" class="td2_b">
                                <asp:RadioButtonList ID="RBL_OPTYPE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" AutoPostBack="true">
                                    <asp:ListItem Value="1" Text="核准入會資料" ></asp:ListItem>
                                    <asp:ListItem Value="2" Text="不核准入會資料查詢" ></asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>        
                    </asp:PlaceHolder>

                    <asp:PlaceHolder ID="PLH_OPTYPE_1" runat="server" Visible="false">
                        <table class="CRTable" width="100%" Cellspacing="0">
                            <tr>
                                <td width="10%" class="th1c_b">
                                    點選
                                </td>
                                <td width="20%" class="th1c_b">
                                    姓名
                                </td>
                                <td width="40%" class="th1c_b">
                                    通訊地址
                                </td>
                                <td width="15%" class="th1c_b">
                                    手機
                                </td>
                                <td width="15%" class="th1c_b">
                                    e-mail
                                </td>
                            </tr>
                            <asp:Repeater ID="RP_APV" runat="server">
                                <ItemTemplate>
                                    <tr>
                                        <td width="10%" class="td2c_b">
                                            <!--點選-->
                                            <asp:RadioButton ID="RB_CHOOSE" runat="server" AutoPostBack="true" OnCheckedChanged="RB_CHOOSE_OnCheckedChanged" />
                                            <input type="hidden" id="MB_DATE" runat="server" value='<%#Container.DataItem("MB_DATE")%>' />
                                        </td>
                                        <td width="20%" class="td2c_b">
                                            <!--姓名-->
                                            <%#Container.DataItem("MB_NAME")%>
                                        </td>
                                        <td width="40%" class="td2c_b">
                                            <!--通訊地址-->
                                            <%#Container.DataItem("MB_CITY")%><%#Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
                                        </td>
                                        <td width="15%" class="td2c_b">
                                            <!--手機-->
                                            <%#Container.DataItem("MB_MOBIL")%>
                                        </td>
                                        <td width="15%" class="td2c_b">
                                            <!--e-mail-->
                                            <%#Container.DataItem("MB_EMAIL")%>
                                        </td>
                                    </tr>
                                </ItemTemplate>
                            </asp:Repeater>
                        </table>

                        <asp:PlaceHolder ID="PLH_MB_MEMBER" runat="server" Visible=false >
                            <table class="CRTable_Top" width="100%" cellspacing="0">
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;法名/姓名
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_NAME" runat="server"  />
                                        <input type="hidden" id="MB_DATE" runat="server" />
                                    </td>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;性別
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_SEX" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        出生年月日
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:literal ID="MB_BIRTH" runat="server"  />
                                    </td>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;身分別
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:CheckBoxList ID="MB_IDENTIFY" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Text="會員" Value="A"></asp:ListItem>
                                            <asp:ListItem Text="委員" Value="B"></asp:ListItem>
                                            <asp:ListItem Text="學員" Value="C"></asp:ListItem>
                                            <asp:ListItem Text="法工" Value="D"></asp:ListItem>
                                        </asp:CheckBoxList>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;手機
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_MOBIL" runat="server"  />
                                    </td>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;電話
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_TEL_Pre" runat="server"  />
                                        ─
                                        <asp:Literal ID="MB_TEL" runat="server"  />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        e-mail
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_EMAIL" runat="server" />
                                    </td>
                                    <td width="15%" class="th1c_b">
                                        身分證字號
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_ID" runat="server"  />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        學歷
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_EDU" runat="server" />
                                    </td>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;所屬區
                                    </td>
                                    <td width="35%" class="td2_b">
                                        <asp:Literal ID="MB_AREA" runat="server" />
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;委員
                                        <asp:Literal ID="MB_LEADER" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;通訊地址
                                    </td>
                                    <td width="85%" class="td2_b" colspan="3">
                                        <asp:Literal ID="MB_CITY" runat="server" />
                                        <asp:Literal ID="MB_VLG" runat="server" />
                                        <asp:Literal ID="MB_ADDR" runat="server" />
                                        郵遞區號 :
                                        <asp:Literal ID="MB_ZIP" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;戶籍地址
                                    </td>
                                    <td width="85%" class="td2_b" colspan="3">
                                        <asp:Literal ID="MB_CITY1" runat="server" />
                                        <asp:Literal ID="MB_VLG1" runat="server" />
                                        <asp:Literal ID="MB_ADDR2" runat="server" />
                                        郵遞區號 :
                                        <asp:Literal ID="MB_ZIP1" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        會員類別
                                    </td>
                                    <td width="85%" class="td2_b" colspan="3">
                                        <asp:Literal ID="MB_FAMILY" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="15%" class="th1c_b">
                                        <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;護持會員
                                    </td>
                                    <td width="85%" class="td2_b" colspan="3">
                                        <asp:CheckBoxList ID="VIP" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                                            <asp:ListItem Value="1" Text="長期護持會員" />
                                            <asp:ListItem Value="2" Text="種子護法" />
                                        </asp:CheckBoxList>
                                    </td>
                                </tr>
                            </table>
                            <div align="center">
                                <asp:Button ID="btn_Y" runat="server" CssClass="bt" Text="核准入會" />
                                &nbsp;
                                <asp:Button ID="btn_D" runat="server" CssClass="bt" Text="不核准入會" />
                            </div>
                        </asp:PlaceHolder>
                    </asp:PlaceHolder>

                    <asp:PlaceHolder ID="PLH_OPTYPE_2" runat="server" Visible="false">
                        <table class="CRTable" width="100%" Cellspacing="0">
                            <tr>
                                <td width="10%" class="th1c_b">
                                    點選
                                </td>
                                <td width="20%" class="th1c_b">
                                    姓名
                                </td>
                                <td width="30%" class="th1c_b">
                                    通訊地址
                                </td>
                                <td width="15%" class="th1c_b">
                                    手機
                                </td>
                                <td width="15%" class="th1c_b">
                                    e-mail
                                </td>
                                <td width="10%" class="th1c_b">
                                    刪除
                                </td>
                            </tr>
                            <asp:Repeater ID="RP_APV_D" runat="server">
                                <ItemTemplate>
                                    <tr>
                                        <td width="10%" class="td2c_b">
                                            <!--點選-->
                                            <asp:Button ID="btnCANCEL" runat="server" CssClass="bt" Text="取消不核准" CommandArgument='<%#Container.DataItem("MB_DATE")%>' CommandName="CANCEL" />
                                        </td>
                                        <td width="20%" class="td2c_b">
                                            <!--姓名-->
                                            <%#Container.DataItem("MB_NAME")%>
                                        </td>
                                        <td width="30%" class="td2c_b">
                                            <!--通訊地址-->
                                            <%#Container.DataItem("MB_CITY")%><%#Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
                                        </td>
                                        <td width="15%" class="td2c_b">
                                            <!--手機-->
                                            <%#Container.DataItem("MB_MOBIL")%>
                                        </td>
                                        <td width="15%" class="td2c_b">
                                            <!--e-mail-->
                                            <%#Container.DataItem("MB_EMAIL")%>
                                        </td>
                                        <td width="10%" class="td2c_b">
                                            <asp:ImageButton ID="IMG_DEL" runat="server" CommandArgument='<%#Container.DataItem("MB_DATE")%>' CommandName="DEL" />
                                        </td>
                                    </tr>
                                </ItemTemplate>
                            </asp:Repeater>
                        </table>
                    </asp:PlaceHolder>
                    <!--顯示下方外框線-->
                    <!--#include virtual="~/inc/MBSCTableEnd.inc"-->     
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
