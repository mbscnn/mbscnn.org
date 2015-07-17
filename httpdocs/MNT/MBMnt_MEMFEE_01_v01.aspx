<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_MEMFEE_01_v01.aspx.vb" Inherits="MBSC.MBMnt_MEMFEE_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
    <title>繳款入帳/沖帳覆核作業</title>
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
                    <asp:Literal ID="LTL_TITLE" runat="server" Text="繳款覆核修正/入帳/沖帳作業" />            
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <asp:PlaceHolder ID="PLH_QRY" runat="server">        
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="100%" class="th1_b">
                            請先輸入
                        </td>
                    </tr>
                    <asp:PlaceHolder ID="PLH_QRY_SEL" runat="server" >
                    <tr>
                        <td width="100%" class="td2_b">
                            <asp:RadioButton ID="RB_QRY_1" runat="server" GroupName="RB_QRY" />
                            覆核修正(繳款作業第2關)
                            <asp:RadioButton ID="RB_QRY_1_2" runat="server" GroupName="RB_QRY" />
                            過帳/列印傳票(繳款作業第3關)
                            <asp:RadioButton ID="RB_QRY_2" runat="server" GroupName="RB_QRY" />
                            已入帳(可沖帳或查詢已入帳)
                            <asp:Button ID="bt_Qry_SEL" runat="server" CssClass="bt" Text="確　　定" />
                        </td>
                    </tr>
                    </asp:PlaceHolder>
                    <asp:PlaceHolder ID="PLH_QRY_SEL_SUB_1" runat="server" Visible=false  >
                        <tr>
                            <td width="100%" class="td2_b">
                                <asp:RadioButton ID="RP_QRY_SUB_1_1" runat="server" GroupName="RP_QRY_SUB_1" />
                                所屬區:
                                <asp:DropDownList ID="MB_AREA_SUB_1" runat="server" />
                                <asp:RadioButton ID="RP_QRY_SUB_1_2" runat="server" GroupName="RP_QRY_SUB_1" />
                                繳款年月：
                                西元
                                <asp:TextBox ID="txt_QRY_SUB_1_YYY" runat="server" CssClass="bordernum" MaxLength=4 Columns=4 />
                                年
                                <asp:TextBox ID="txt_QRY_SUB_1_MM" runat="server" CssClass="bordernum" MaxLength=2 Columns=2 />
                                月
                                <asp:Button ID="bt_Qry_SUB_1" runat="server" CssClass="bt" Text="確　　定" />
                                <asp:Button ID="bt_Cancel_SUB_1" runat="server" CssClass="bt" Text="取　　消" />
                            </td>
                        </tr>
                    </asp:PlaceHolder>
                    <asp:PlaceHolder ID="PLH_QRY_SEL_SUB_2" runat="server" Visible=false  >
                        <tr>
                            <td width="100%" class="td2_b">
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
                                <asp:Button ID="bt_Qry_SUB_2" runat="server" CssClass="bt" Text="確　　定" />
                                <asp:Button ID="bt_Cancel_SUB_2" runat="server" CssClass="bt" Text="取　　消" />
                            </td>
                        </tr>
                    </asp:PlaceHolder>
                </table>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMREV" runat="server" Visible=false >
                <asp:PlaceHolder ID="PLH_MB_MEMREV_DATA" runat="server" Visible=false >
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
                                <asp:PlaceHolder ID="PLH_MB_TX_DATE" runat="server" >
                                <asp:TextBox ID="MB_TX_DATE_YYY" runat="server" CssClass="bordernum" Columns=4 MaxLength=4 />
                                年
                                <asp:TextBox ID="MB_TX_DATE_MM" runat="server" CssClass="bordernum" Columns=2 MaxLength=2 />
                                月
                                <asp:TextBox ID="MB_TX_DATE_DD" runat="server" CssClass="bordernum" Columns=2 MaxLength=2 />
                                日
                                </asp:PlaceHolder>
                            </td>
                            <td width="15%" class="th1c_b">
                                功德項目
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:Literal ID="MB_ITEMID" runat="server" />
                                <asp:DropDownList ID="DDL_MB_ITEMID" runat="server" AutoPostBack="true" />
                                <asp:PlaceHolder ID="PLH_MB_MEMTYP" runat="server" Visible=false >
                                &nbsp;
                                會員類別：
                                <asp:Literal ID="MB_MEMTYP" runat="server" />
                                <asp:RadioButtonList ID="RBL_MB_MEMTYP" runat="server" RepeatDirection=Horizontal RepeatLayout=Flow AutoPostBack="true" />
                                </asp:PlaceHolder>
                                <asp:PlaceHolder ID="PLH_PROJCODE" runat="server" Visible="false">
                                    專案名稱：
                                    <asp:DropDownList ID="PROJCODE" runat="server" />
                                </asp:PlaceHolder>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="th1c_b">
                                繳款方式
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:Literal ID="MB_FEETYPE" runat="server" />
                                <asp:RadioButtonList ID="RBL_MB_FEETYPE" runat="server" RepeatDirection=Horizontal RepeatLayout=Flow AutoPostBack="true" />
                            </td>
                            <td width="15%" class="th1c_b">
                                繳款金額
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:Literal ID="MB_TOTFEE" runat="server" />
                                <asp:TextBox ID="TXT_MB_TOTFEE" runat="server" CssClass="bordernum" Columns="7" MaxLength=9 />

                                <asp:PlaceHolder ID="PLH_MB_TOTFEE_MM" runat="server" Visible=false>
                                    每月金額：
                                    <asp:Literal ID="MB_TOTFEE_MM" runat="server" />
                                    <asp:TextBox ID="TXT_MB_TOTFEE_MM" runat="server" CssClass="bordernum" Columns="7" MaxLength=9 AutoPostBack="true" />
                                </asp:PlaceHolder>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="th1c_b">
                                繳款分配項目
                            </td>
                            <td width="85%" class="td2_b" colspan="3" >
                                <asp:CheckBox ID="FUND1" runat="server" Text="護僧基金" />
                                <asp:CheckBox ID="FUND2" runat="server" Text="建設基金" />
                                <asp:CheckBox ID="FUND3" runat="server" Text="弘法基金" />
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="th1c_b">
                                收據捐款名稱
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:Literal ID="MB_RECNAME" runat="server" />
                                <asp:TextBox ID="TXT_MB_RECNAME" runat="server" CssClass="bordertxt" />
                            </td>
                            <td width="15%" class="th1c_b">
                                所屬區
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:Literal ID="MB_AREA" runat="server" />
                                <asp:DropDownList ID="DDL_MB_AREA" runat="server" AutoPostBack="true" />
                                &nbsp;
                                委員：
                                <asp:Literal ID="MB_LEADER" runat="server" />
                                <asp:DropDownList ID="DDL_MB_LEADER" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="th1c_b">
                                會費期間
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:PlaceHolder ID="PLH_SHOW_MB_MEMFEE" runat="server">
                                    <asp:Literal ID="MB_MEMFEE" runat="server" />
                                </asp:PlaceHolder>                        
                        
                                <asp:PlaceHolder ID="PLH_MB_MEMFEE" runat="server">
                                <asp:TextBox ID="MB_MEMFEE_SY" runat="server" CssClass="bordernum" Columns="4" />年
                                <asp:TextBox ID="MB_MEMFEE_SM" runat="server" CssClass="bordernum" Columns="3" />月
                                至
                                <asp:TextBox ID="MB_MEMFEE_EY" runat="server" CssClass="bordernum" Columns="4" />年
                                <asp:TextBox ID="MB_MEMFEE_EM" runat="server" CssClass="bordernum" Columns="3" />月
                                </asp:PlaceHolder>

                                <asp:Literal ID="MB_DESC" runat="server" Visible=false />
                            </td>
                            <td width="15%" class="th1c_b">
                                付款方式
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:Literal ID="MB_PAY_TYPE" runat="server" />
                                <asp:RadioButtonList ID="RBL_MB_PAY_TYPE" runat="server" RepeatDirection=Horizontal RepeatLayout=Flow  />
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="th1c_b">
                                是否開立收據
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:RadioButtonList ID="MB_PRINT" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" AutoPostBack="true" >
                                    <asp:ListItem Text="是" Value="1" />
                                    <asp:ListItem Text="否" Value="0" />
                                </asp:RadioButtonList>
                            </td>
                            <td width="15%" class="th1c_b">
                             <asp:Literal ID="LTL_MB_SEND_PRINT" runat="server"  Text="給收據方式" Visible="false" />
                            </td>
                            <td width="35%" class="td2_b">
                                <asp:DropDownList ID="MB_SEND_PRINT" runat="server" Visible="false" />
                            </td>
                        </tr>
                        <asp:PlaceHolder ID="PLH_MB_PAY_TYPE" runat="server" Visible=false >
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
                        </asp:PlaceHolder>
                    </table>
                    <div align=center >
                        <asp:Button ID="bt_Cancel_Data" runat="server" CssClass="bt" Text="取　　消" />
                        &nbsp;
                        <asp:Button ID="btnModify" runat="server" CssClass="bt" Text="覆核修正" Visible=false />
                        &nbsp;
                        <asp:Button ID="btnPrint" runat="server" CssClass="bt" Text="列印傳票" Visible=false />
                        &nbsp;
                        <asp:Button ID="bt_Approve_Data" runat="server" CssClass="bt" Text="過帳" Visible=false />
                        &nbsp;
                        <asp:Button ID="bt_Acct_Data" runat="server" CssClass="bt" Text="沖帳" Visible=false />
                    </div>
                </asp:PlaceHolder>

                <asp:PlaceHolder ID="PLH_MB_MEMREV_LIST" runat="server" >
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
                        <td width="10%" class="th1c_b" id="TD_PRINT" runat="server" visible=false >
                            收據
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
                                <td width="10%" class="td2c_b" id="TD_PRINT" runat=server visible=false>
                                    <!--收據-->
                                    <%#get_Print_Desc(Container.DataItem("MB_PRINT"),Container.DataItem("MB_REBREV"))%>
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
            </asp:PlaceHolder>
            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"-->
        </td>
    </tr>
</table>
    </form>
</body>
</html>
