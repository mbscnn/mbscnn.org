<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Print_01_v01.aspx.vb" Inherits="MBSC.MBMnt_Print_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
    <title>列印回收收據作業</title>
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
                    社團法人台灣佛陀原始正法學會
                    <br />
                    列印/回收收據作業
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <asp:PlaceHolder ID="PLH_QRY" runat="server">
            <table class="CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td class="th1_b" width="100%">
                        請先選擇
                    </td>
                </tr>
                <tr>
                    <td class="td2_b" width="100%">
                        <asp:RadioButtonList ID="RBL_Choose" runat="server" RepeatDirection=Vertical RepeatLayout=Flow>
                            <asp:ListItem Text="列印收據(已覆核過帳且未列印收據者)" Value="0"></asp:ListItem>
                            <asp:ListItem Text="回收收據 (已列印收據者，回收銷毀，才可退款作沖帳)" Value="1"></asp:ListItem>
                            <asp:ListItem Text="補印收據 (已列印收據者)" Value="2"></asp:ListItem>
                        </asp:RadioButtonList> 
                    </td>
                </tr>
                <tr>
                    <td class="td2_b" width="100%">
                        <asp:RadioButton ID="RP_QRY_SUB_2_1" runat="server" GroupName="RP_QRY_SUB_2" />
                        所屬區:
                        <asp:DropDownList ID="MB_AREA_SUB_2" runat="server" />
                        <asp:RadioButton ID="RP_QRY_SUB_2_2" runat="server" GroupName="RP_QRY_SUB_2" />
                        繳款年月：
                        西元
                        <asp:TextBox ID="txt_QRY_SUB_2_BEGYYY" runat="server" CssClass="bordernum" MaxLength=4 Columns=4 />
                        年
                        <asp:TextBox ID="txt_QRY_SUB_2_BEGMM" runat="server" CssClass="bordernum" MaxLength=2 Columns=2 />
                        月
                        ~
                        西元
                        <asp:TextBox ID="txt_QRY_SUB_2_ENDYYY" runat="server" CssClass="bordernum" MaxLength=4 Columns=4 />
                        年
                        <asp:TextBox ID="txt_QRY_SUB_2_ENDMM" runat="server" CssClass="bordernum" MaxLength=2 Columns=2 />
                        月
                        <asp:RadioButton ID="RP_QRY_SUB_2_3" runat="server" GroupName="RP_QRY_SUB_2" />
                        姓名
                        <asp:TextBox ID="txt_QRY_MB_NAME" runat="server" CssClass="bordertxt" />                
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Button ID="bt_Qry_SUB_2" runat="server" CssClass="bt" Text="確　　定" />
            </div>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMREV" runat="server" Visible=false >
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            法名/姓名
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_NAME" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_MEMSEQ" runat="server" />
                            <input type="hidden" id="HID_MB_MEMSEQ" runat="server" />
                            <input type="hidden" id="HID_MB_SEQNO" runat="server" />
                            <input type="hidden" id="HID_MB_ITEMID" runat="server" />
                            <input type="hidden" id="HID_MB_MEMFEE_SY" runat="server" />
                            <input type="hidden" id="HID_MB_MEMFEE_EY" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            通訊地址
                        </td>
                        <td width="85%" class="td2_b" colspan=3 >
                            <asp:Literal ID="MB_ADDR" runat="server" />
                        </td>                
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            繳款日期
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_TX_DATE" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            功德項目
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_ITEMID" runat="server" />
                            <asp:PlaceHolder ID="PLH_MB_MEMTYP" runat="server" Visible=false >
                            &nbsp;
                            會員類別：
                            <asp:Literal ID="MB_MEMTYP" runat="server" />
                            </asp:PlaceHolder>
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            繳款方式
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_FEETYPE" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            繳款金額
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_TOTFEE" runat="server" />
                            <asp:PlaceHolder ID="PLH_MB_TOTFEE_MM" runat="server" Visible=false>
                                每月金額：
                                <asp:Literal ID="MB_TOTFEE_MM" runat="server" />
                            </asp:PlaceHolder>
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            收據捐款名稱
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_RECNAME" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            所屬區
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_AREA" runat="server" />
                            &nbsp;
                            委員：
                            <asp:Literal ID="MB_LEADER" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            會費期間
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_MEMFEE" runat="server" />
                            <asp:Literal ID="MB_DESC" runat="server" Visible=false />
                        </td>
                        <td width="15%" class="th1c_b">
                            付款方式
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_PAY_TYPE" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            票據到期日
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="NOTE_DUE_DATE" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            票據號碼
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="NOTE_NO" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            發票行
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="NOTE_BANK" runat="server" />
                            銀行
                            <asp:Literal ID="NOTE_BR" runat="server" />
                            分行
                        </td>
                        <td width="15%" class="th1c_b">
                            發票人
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="NOTE_HOLDER" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            票據金額
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="NOTE_AMT" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            補印次數
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="MB_REISU" runat="server" />
                        </td>
                    </tr>
                </table>
                <div align=center >
                    <asp:Button ID="bt_Cancel_Data" runat="server" CssClass="bt" Text="取　　消" />
                    <asp:Button ID="bt_Print" runat="server" CssClass="bt" Text="列印收據" Visible=false />
                    <asp:Button ID="bt_Return" runat="server" CssClass="bt" Text="回收收據" Visible=false />
                    <asp:Button ID="bt_Append" runat="server" CssClass="bt" Text="補印收據" Visible=false />
                </div>        
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMREV_LIST" runat="server" Visible=false >
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            姓名
                        </td>
                        <td width="10%" class="th1c_b">
                            繳款日期
                        </td>
                        <td width="10%" class="th1c_b">
                            功德項目
                        </td>
                        <td width="10%" class="th1c_b">
                            繳款金額
                        </td>
                        <td width="15%" class="th1c_b">
                            收據捐款名稱
                        </td>
                        <td width="10%" class="th1c_b">
                            所屬區
                        </td>
                        <td width="10%" class="th1c_b">
                            委員
                        </td>
                        <td width="10%" class="th1c_b">
                            補印次數
                        </td>
                        <td width="10%" class="th1c_b">
                            檢視
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_MB_MEMREV" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="15%" class="td2c_b">
                                    <!--姓名-->
                                    <%#Container.DataItem("MB_NAME")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--繳款日期-->
                                    <%#get_DECIMAL_DATE(Container.DataItem("MB_TX_DATE"))%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--功德項目-->
                                    <%#get_MB_ITEMID_TEXT(Container.DataItem("MB_ITEMID"))%>
                                </td>
                                <td width="10%" class="td2r_b">
                                    <!--繳款金額-->
                                    <%#get_MB_TOTFEE_TEXT(Container.DataItem("MB_TOTFEE"))%>
                                </td>
                                <td width="15%" class="td2c_b">
                                    <!--收據捐款名稱-->
                                    <%#Container.DataItem("MB_RECNAME")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--所屬區-->
                                    <%#getMB_AREA(Container.DataItem("MB_AREA"))%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--委員-->
                                    <%#Container.DataItem("MB_LEADER")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--補印次數-->
                                    <%#Container.DataItem("MB_REISU")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--檢視-->
                                    <asp:ImageButton ID="IMG_EDIT" runat="server" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' CommandName="EDIT" ImageUrl="~/img/search2.gif" />
                                    <input type="hidden" id="HID_MB_SEQNO" runat="server" value='<%#Container.DataItem("MB_SEQNO")%>' />                            
                                </td>
                            </tr>                    
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
                <div align="center">
                    <asp:Button ID="bt_Back" runat="server" CssClass="bt" Text="回選擇畫面" />
                </div>
            </asp:PlaceHolder>
            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"-->     
        </td>
    </tr>
</table>   
    </form>
</body>
</html>
