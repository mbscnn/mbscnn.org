<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_TYPE_01_v01.aspx.vb" Inherits="MBSC.MBQry_TYPE_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->

</head>
<body>
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
<table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
    <tr>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
            <!-- #include virtual="~/inc/Signin.inc" -->
            <div align="center">
                <h2 style="color: #800080; font-family: 標楷體">
                    MBSC各類報表查詢
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <asp:PlaceHolder ID="PLH_QTYPE" runat="server" >
            <table width="100%" class="CRTable_Top" cellspacing="0">
                <tr>
                    <td width="15%" class="th1c_b">
                        報表查詢格式
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_TYPE" runat="server" >
                            <asp:ListItem Text="請選擇" Value="" />
                            <asp:ListItem Text="學員資料查詢" Value="1" />
                            <asp:ListItem Text="會員繳款資料查詢" Value="2" />
                            <asp:ListItem Text="課程報名查詢" Value="3" />
                            <asp:ListItem Text="會員收入查詢" Value="4" />
                            <asp:ListItem Text="一般會員繳款年報表" Value="5" />
                            <asp:ListItem Text="種籽會員繳款報表" Value="6" />
                        </asp:DropDownList>
                        <asp:Button ID="btnQTYPE" runat="server" CssClass="bt" Text="確定" />
                    </td>
                </tr>
            </table>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_PARAS" runat="server" Visible="false">
            <table width="100%" class="CRTable_Top" cellspacing="0">
                <tr>
                    <td width="15%" class="th1c_b">
                        報表查詢格式
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:Literal ID="LTL_TYPE" runat="server" />
                    </td>
                </tr>
                <tr id="TR_MB_MEMTYP" runat="server" visible="false">
                    <td width="15%" class="th1c_b">
                        繳款會員種類
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="MB_MEMTYP" runat="server" />
                        繳款方式:
                        <asp:DropDownList ID="MB_FEETYPE" runat="server" />
                    </td>
                </tr>

                <asp:PlaceHolder ID="PLH_3" runat="server" visible="false">
                <tr>
                    <td width="15%" class="th1c_b">
                      課程類別
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="ddl_CLASS_TYPE" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程地點
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="ddl_PLACE" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程開始年月
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:DropDownList ID="ddl_StartYear" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="ddl_StartMonth" runat="server">
                        </asp:DropDownList>月至
                        <asp:DropDownList ID="ddl_EndYear" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="ddl_EndMonth" runat="server">
                        </asp:DropDownList>月
                        <asp:Button ID="btnQRYClass" runat="server" CssClass="bt" Text="查詢課程清單" Visible="false" />
                    </td>
                </tr>
                </asp:PlaceHolder>
                <tr id="TR_BE_YYYMMDD" runat="server">
                    <td width="15%" class="th1c_b">
                        <asp:Literal ID="LTL_PERIOD" runat="server" Text="輸入起訖日期" />
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:TextBox ID="TXT_BY" runat="server" MaxLength="4" CssClass="bordernum" Columns="4" />年(YYYY)
                        <asp:TextBox ID="TXT_BM" runat="server" MaxLength="2" CssClass="bordernum" Columns="2" />月
                        <asp:TextBox ID="TXT_BD" runat="server" MaxLength="2" CssClass="bordernum" Columns="2" />日
                        &nbsp;至&nbsp;
                        <asp:TextBox ID="TXT_EY" runat="server" MaxLength="4" CssClass="bordernum" Columns="4" />年(YYYY)
                        <asp:TextBox ID="TXT_EM" runat="server" MaxLength="2" CssClass="bordernum" Columns="2" />月
                        <asp:TextBox ID="TXT_ED" runat="server" MaxLength="2" CssClass="bordernum" Columns="2" />日                       
                    </td>
                </tr>
                <tr id="TR_BE_YYY" runat="server" visible="false">
                    <td width="15%" class="th1c_b">
                        查詢起訖年度
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:TextBox ID="TXT_B_YYY" runat="server" MaxLength="4" CssClass="bordernum" Columns="4" />年(YYYY)
                        &nbsp;至&nbsp;
                        <asp:TextBox ID="TXT_E_YYY" runat="server" MaxLength="4" CssClass="bordernum" Columns="4" />年(YYYY)
                    </td>
                </tr>
                <asp:PlaceHolder ID="PLH_MB_AREA" runat="server" Visible="false">
                    <tr>
                        <td width="15%" class="th1c_b">
                            所屬區
                        </td>
                        <td width="85%" class="td2_b">
                            <asp:DropDownList ID="DDL_MB_AREA" runat="server" AutoPostBack="true" />
                            委員 : 
                            <asp:DropDownList ID="DDL_MB_LEADER" runat="server" >
                                <asp:ListItem Text="所有委員" Value="" ></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                </asp:PlaceHolder>
                <asp:PlaceHolder ID="PLH_MB_CLASS" runat="server" Visible="false" >
                    <tr>
                        <td width="15%" class="th1c_b">
                            課程編號
                        </td>
                        <td width="85%" class="td2_b">
                            <asp:DropDownList ID="DDL_MB_CLASS" runat="server" />
                        </td>
                    </tr>
                </asp:PlaceHolder>
            </table>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_DATA" runat="server" Visible="false">
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
            <asp:PlaceHolder ID="PLH_DATA_1" runat="server" Visible="false">
            <table class="CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="10%" class="th1c_b">
                        所屬區
                    </td>
                    <td width="15%" class="th1c_b">
                        委員
                    </td>
                    <td width="15%" class="th1c_b">
                        會員姓名
                    </td>
                    <td width="15%" class="th1c_b">
                        會員編號
                    </td>
                    <td width="15%" class="th1c_b">
                        會員類別
                    </td>
                    <td width="15%" class="th1c_b">
                        電話
                    </td>
                    <td width="15%" class="th1c_b">
                        手機
                    </td>
                </tr>
                <asp:Repeater ID="RP_1" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td width="10%" class="td2c_b">
                                <!--所屬區-->
                                <%#getMB_AREA(Container.DataItem("MB_AREA"))%>
                            </td>
                            <td width="15%" class="td2c_b">
                                <!--委員-->
                                <%#Container.DataItem("MB_LEADER")%>
                            </td>
                            <td width="15%" class="td2c_b">
                                <!--會員姓名-->
                                <%#Container.DataItem("MB_NAME")%>
                            </td>
                            <td width="15%" class="td2c_b">
                                <!--會員編號-->
                                <asp:LinkButton ID="LB_MB_MEMSEQ" runat="server" Text='<%#getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>' CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' />
                            </td>
                            <td width="15%" class="td2c_b">
                                <!--會員類別-->
                                <%#getMB_FAMILY(Container.DataItem("MB_FAMILY"))%>
                            </td>
                            <td width="15%" class="td2c_b">
                                <!--電話-->
                                <%#Container.DataItem("MB_TEL")%>
                            </td>
                            <td width="15%" class="td2c_b">
                                <!--手機-->
                                <%#Container.DataItem("MB_MOBIL")%>
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_DATA_2" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="10%" class="th1c_b">
                            繳款日期
                        </td>
                        <td width="10%" class="th1c_b">
                            所屬區
                        </td>
                        <td width="15%" class="th1c_b">
                            委員
                        </td>
                        <td width="15%" class="th1c_b">
                            會員姓名
                        </td>
                        <td width="10%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="15%" class="th1c_b">
                            功德項目
                        </td>
                        <td width="15%" class="th1c_b">
                            繳款金額
                        </td>
                        <td width="10%" class="th1c_b">
                            筆數
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_2" runat="server">
                        <ItemTemplate>
                            <INPUT type="hidden" id="HID_MB_MEMSEQ" runat="server" value='<%#Container.DataItem("MB_MEMSEQ")%>' />
                            <asp:Repeater ID="RP_2_DTL" runat="server" OnItemDataBound="RP_2_DTL_OnItemDataBound">
                                <ItemTemplate>
                                    <tr>
                                        <td id="TD_1" runat="server" width="10%" class="td2c_b">
                                            <!--繳款日期-->
                                            <asp:Literal ID="LTL_MB_TX_DATE" runat="server" Text='<%#Container.DataItem("MB_TX_DATE")%>' />
                                        </td>
                                        <td id="TD_2" runat="server" width="10%" class="td2c_b">
                                            <!--所屬區-->
                                            <%#getMB_AREA(Container.DataItem("MB_AREA"))%>
                                        </td>
                                        <td id="TD_3" runat="server" width="15%" class="td2c_b">
                                            <!--委員-->
                                            <%#Container.DataItem("MB_LEADER")%>
                                        </td>
                                        <td id="TD_4" runat="server" width="15%" class="td2c_b">
                                            <!--會員姓名-->
                                            <%#Container.DataItem("MB_NAME")%>
                                        </td>
                                        <td id="TD_5" runat="server" width="10%" class="td2c_b">
                                            <!--會員編號-->
                                            <%#getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
                                        </td>
                                        <td id="TD_6" runat="server" width="15%" class="td2c_b">
                                            <!--功德項目-->
                                            <%#getMB_ITEMID(Container.DataItem("MB_ITEMID"))%>
                                        </td>
                                        <td id="TD_7" runat="server" width="15%" class="td2r_b">
                                            <!--繳款金額-->
                                            <asp:Literal ID="LTL_MB_TOTFEE" runat="server" />
                                        </td>
                                        <td id="TD_8" runat="server" width="10%" class="td2c_b">
                                            <!--筆數-->
                                            <%#Container.DataItem("CNT")%>
                                        </td>
                                    </tr>
                                </ItemTemplate>
                            </asp:Repeater>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_DATA_3" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="10%" class="th1c_b">
                            報課會員
                        </td>
                        <td width="10%" class="th1c_b">
                            性別
                        </td>
                        <td width="10%" class="th1c_b">
                            會員信箱
                        </td>
                        <td width="10%" class="th1c_b">
                            會員電話
                        </td>
                        <td width="10%" class="th1c_b">
                            會員手機
                        </td>
                        <td width="10%" class="th1c_b">
                            課程編號
                        </td>
                        <td width="10%" class="th1c_b">
                            課程名稱
                        </td>
                        <td width="10%" class="th1c_b">
                            錄取
                        </td>
                        <td width="10%" class="th1c_b">
                            回信出席
                        </td>
                        <td width="10%" class="th1c_b">
                            未出席
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_3" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="10%" class="td2c_b">
                                    <!--報課會員-->
                                    <asp:LinkButton ID="LBL_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>' CommandName="SIGN" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                    <input type="hidden" id="HID_MB_BATCH" runat="server" value='<%#Container.DataItem("MB_BATCH")%>' />
                                    <input type="hidden" id="HID_MB_MEMSEQ" runat="server" value='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                    <input type="hidden" id="HID_MB_SEQ" runat="server" value='<%#Container.DataItem("MB_SEQ")%>' />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--性別-->
                                    <asp:Literal ID="LTL_MB_SEX" runat="server" />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--會員信箱-->
                                    <%#Container.DataItem("MB_EMAIL")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--會員電話-->
                                    <%#Container.DataItem("MB_TEL")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--會員手機-->
                                    <%#Container.DataItem("MB_MOBIL")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--課程編號-->
                                    <asp:Literal ID="LTL_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />                                
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--課程名稱-->
                                    <asp:Literal ID="MB_CLASS_NAME" runat="server" />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--錄取-->
                                    <%#getMB_CHKFLAG(Container.DataItem("MB_SEQ"), Container.DataItem("MB_BATCH"), Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_FWMK"), Container.DataItem("MB_SORTNO"), Container.DataItem("MB_CHKFLAG"), Container.DataItem("MB_CDATE"), Container.DataItem("MB_CREDATETIME"))%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--回信出席-->
                                    <asp:Literal ID="LTL_MB_RESP" runat="server" />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--未出席-->
                                    <asp:CheckBox ID="MB_FWMK" runat="server" />
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_DATA_4" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="7%" class="th1c_b">
                            區別
                        </td>
                        <td width="7%" class="th1c_b">
                            委員
                        </td>
                        <td width="7%" class="th1c_b">
                            日期
                        </td>
                        <td width="7%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="7%" class="th1c_b">
                            捐款人
                        </td>
                        <td width="7%" class="th1c_b">
                            會費起訖日
                        </td>
                        <td width="7%" class="th1c_b">
                            會員種類
                        </td>
                        <td width="7%" class="th1c_b">
                            繳款方式
                        </td>
                        <td width="7%" class="th1c_b">
                            每月金額
                        </td>
                        <td width="7%" class="th1c_b">
                            繳款金額
                        </td>
                        <td width="4%" class="th1c_b">
                            護僧基金
                        </td>
                        <td width="4%" class="th1c_b">
                            建設基金
                        </td>
                        <td width="4%" class="th1c_b">
                            弘法基金
                        </td>
                        <td width="7%" class="th1c_b">
                            列印收據
                        </td>
                        <td width="7%" class="th1c_b">
                            收據編號
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_4" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="7%" class="td2c_b">
                                    <!--區別-->
                                    <asp:Literal ID="LTL_MB_AREA" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--委員-->
                                    <asp:Literal ID="LTL_MB_LEADER" runat="server" Text='<%#Container.DataItem("MB_LEADER")%>' />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--日期-->
                                    <asp:Literal ID="LTL_MB_TX_DATE" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--會員編號-->
                                    <asp:Literal ID="LTL_MB_MEMSEQ" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--捐款人-->
                                    <asp:Literal ID="LTL_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>' />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--會費起訖日-->
                                    <asp:Literal ID="LTL_MB_MEMFEE_SYMEYM" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--會員種類-->
                                    <asp:Literal ID="LTL_MB_MEMTYP" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--繳款方式-->
                                    <asp:Literal ID="LTL_MB_FEETYPE" runat="server" />
                                </td>
                                <td width="7%" class="td2r_b">
                                     <!--每月金額-->
                                    <asp:Literal ID="LTL_MB_MONFEE" runat="server" />
                                </td>
                                <td width="7%" class="td2r_b">
                                     <!--繳款金額-->
                                    <asp:Literal ID="LTL_MB_TOTFEE" runat="server" />
                                </td>
                                <td width="4%" class="td2c_b">
                                     <!--護僧基金-->
                                    <asp:Literal ID="LTL_FUND1" runat="server" />
                                </td>
                                <td width="4%" class="td2c_b">
                                     <!--建設基金-->
                                    <asp:Literal ID="LTL_FUND2" runat="server" />
                                </td>
                                <td width="4%" class="td2c_b">
                                     <!--弘法基金-->
                                    <asp:Literal ID="LTL_FUND3" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--列印收據-->
                                    <asp:Literal ID="LTL_MB_PRINT" runat="server" />
                                </td>
                                <td width="7%" class="td2c_b">
                                     <!--收據編號-->
                                    <asp:Literal ID="LTL_MB_RECV" runat="server" />
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_DATA_5" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="8%" class="th1c_b">
                            區別
                        </td>
                        <td width="8%" class="th1c_b">
                            委員
                        </td>
                        <td width="8%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="8%" class="th1c_b">
                            捐款人
                        </td>
                        <td width="8%" class="th1c_b">
                            繳款年度
                        </td>
                        <td width="5%" class="th1c_b">
                            1月
                        </td>
                        <td width="5%" class="th1c_b">
                            2月
                        </td>
                        <td width="5%" class="th1c_b">
                            3月
                        </td>
                        <td width="5%" class="th1c_b">
                            4月
                        </td>
                        <td width="5%" class="th1c_b">
                            5月
                        </td>
                        <td width="5%" class="th1c_b">
                            6月
                        </td>
                        <td width="5%" class="th1c_b">
                            7月
                        </td>
                        <td width="5%" class="th1c_b">
                            8月
                        </td>
                        <td width="5%" class="th1c_b">
                            9月
                        </td>
                        <td width="5%" class="th1c_b">
                            10月
                        </td>
                        <td width="5%" class="th1c_b">
                            11月
                        </td>
                        <td width="5%" class="th1c_b">
                            12月
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_5" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="8%" class="td2c_b">
                                    <!--區別-->
                                    <%#getMB_AREA(Container.DataItem("MB_AREA"))%>
                                </td>
                                <td width="8%" class="td2c_b">
                                    <!--委員-->
                                    <%#Container.DataItem("MB_LEADER")%>
                                </td>
                                <td width="8%" class="td2c_b">
                                    <!--會員編號-->
                                    <%#getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
                                </td>
                                <td width="8%" class="td2c_b">
                                    <!--捐款人-->
                                    <%#Container.DataItem("MB_NAME")%>
                                </td>
                                <td width="8%" class="td2c_b">
                                    <!--繳款年度-->
                                    <%#Container.DataItem("MB_YYYY")%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--1月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M1"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--2月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M2"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--3月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M3"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--4月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M4"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--5月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M5"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--6月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M6"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--7月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M7"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--8月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M8"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--9月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M9"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--10月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M10"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--11月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M11"))%>
                                </td>
                                <td width="5%" class="td2c_b">
                                    <!--12月-->
                                    <%#getFEEAmt(Container.DataItem("MB_M12"))%>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_DATA_6" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            區別
                        </td>
                        <td width="15%" class="th1c_b">
                            委員
                        </td>
                        <td width="10%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="20%" class="th1c_b">
                            捐款人
                        </td>
                        <td width="20%" class="th1c_b">
                            繳款日期
                        </td>
                        <td width="20%" class="th1c_b">
                            繳款金額
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_6" runat="server">
                        <ItemTemplate>
                            <asp:Literal ID="LTL_MB_MEMSEQ" runat="server" Text='<%#Container.DataItem("MB_MEMSEQ")%>' Visible="false" />
                            <asp:Literal ID="LTL_MB_AREA" runat="server" Text='<%#Container.DataItem("MB_AREA")%>' Visible="false" />
                            <asp:Literal ID="LTL_MB_LEADER" runat="server" Text='<%#Container.DataItem("MB_LEADER")%>' Visible="false" />
                            <asp:Literal ID="LTL_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>' Visible="false" />
                            <asp:Repeater ID="RP_6_DTL" runat="server" OnItemDataBound="RP_6_DTL_OnItemDataBound" >
                                <ItemTemplate>
                                    <tr>
                                        <td width="15%" class="td2c_b">
                                            <!--區別-->
                                            <asp:Literal ID="LTL_MB_AREA" runat="server" />
                                        </td>
                                        <td width="15%" class="td2c_b">
                                            <!--委員-->
                                            <%#Container.DataItem("MB_LEADER")%>
                                        </td>
                                        <td width="10%" class="td2c_b">
                                            <!--會員編號-->                                        
                                            <asp:Literal ID="LTL_MB_MEMSEQ" runat="server" text='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                        </td>
                                        <td width="20%" class="td2c_b">
                                            <!--捐款人-->
                                            <%#Container.DataItem("MB_NAME")%>
                                        </td>
                                        <td width="20%" class="td2c_b">
                                            <!--繳款日期-->
                                            <asp:Literal ID="LTL_MB_TX_DATE" runat="server" />
                                        </td>
                                        <td width="20%" class="td2r_b">
                                            <!--繳款金額-->
                                            <asp:Literal ID="LTL_MB_TOTFEE" runat="server" />
                                        </td>
                                    </tr>
                                </ItemTemplate>
                            </asp:Repeater>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </asp:PlaceHolder>
            </asp:PlaceHolder>

            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"-->
            <asp:PlaceHolder ID="PLH_BUTTON" runat="server" Visible="false">
            <div align="center">
                <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="bt" />
                <asp:Button ID="btnMB_FWMK" runat="server" Text="確定" CssClass="bt" />
                <asp:Button ID="btnEXCEL" runat="server" Text="另存EXCEL" CssClass="bt" />
                <asp:Button ID="btnQRY" runat="server" Text="查詢" CssClass="bt" />
            </div>
            </asp:PlaceHolder>
        </td>
    </tr>
</table>
    </form>
</body>
</html>
