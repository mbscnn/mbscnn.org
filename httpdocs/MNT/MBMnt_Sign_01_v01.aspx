<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Sign_01_v01.aspx.vb" Inherits="MBSC.MBMnt_Sign_01_v01" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head runat="server">
    <title></title>
    <script type="text/javascript" language="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/jquery-1.7.1.min.js"%>"></script>
    <script type="text/javascript" language="JavaScript">
        //DataGrid Radiobutton防止多選
        function RadioButtonSelect(aspRadioButtonID) {
            re = new RegExp('rbData')
            for (i = 0; i < document.forms[0].elements.length; i++) {

                elm = document.forms[0].elements[i]

                if (elm.type == 'radio') {
                    if (re.test(elm.name)) {
                        elm.checked = false
                    }
                }
            }
            aspRadioButtonID.checked = true;
        };
        //出家眾
        function CheckMonk(thisobj) {
            if (thisobj == document.getElementById("rbt_MB_MONK_N")) {
                document.getElementById('tr_MB_MONK_1').style.display = 'none';
                document.getElementById('tr_MB_MONK_2').style.display = 'none';
                document.getElementById('tr_MB_MONK_3').style.display = 'none';
                document.getElementById('tr_MB_MONK_4').style.display = 'none';
                return true;
            } else {
                document.getElementById('tr_MB_MONK_1').style.display = '';
                document.getElementById('tr_MB_MONK_2').style.display = '';
                document.getElementById('tr_MB_MONK_3').style.display = '';
                document.getElementById('tr_MB_MONK_4').style.display = '';
                return true;
            }

            return false;
        };
    </script>
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
                    MBSC報名表
                </h2>
            </div>

            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->

            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->

            <table class="CRTable_Top" width="100%" cellspacing="0" runat="server" id="tb_Page1">
                <tr>
                    <td class="th1_b" width="100%" colspan="4">
                        帳號：<asp:Label ID="lbl_USERID" runat="server" Text="" ForeColor="Blue"></asp:Label>
                        <br />
                        <span style="color: #FF0000">(若非報名本人請輸入報名者手機或電話擇一，將選出報名者資料)</span>
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        <asp:RadioButton ID="rbt_Cel" runat="server" Text="手機：" GroupName="SelectType" />
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_Cel" runat="server" Width="45%"></asp:TextBox>
                    </td>
                    <td class="th1c_b" width="15%">
                        <asp:RadioButton ID="rbt_Tel" runat="server" Text="電話：" GroupName="SelectType" />
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_Tel_Zip" runat="server" Width="10%" MaxLength="2"></asp:TextBox>
                        &nbsp;─&nbsp;
                        <asp:TextBox ID="txt_Tel" runat="server" Width="40%" MaxLength="10"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="th1r_b" width="100%" colspan="4" align="right">
                        <asp:Button ID="btn_Qry" runat="server" Text="確定" CssClass="bt" />
                    </td>
                </tr>
            </table>

            <asp:PlaceHolder ID="PLH_MAIL_SAME" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" Cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            點選
                        </td>
                        <td width="20%" class="th1c_b">
                            姓名
                        </td>
                        <td width="35%" class="th1c_b">
                            通訊地址
                        </td>
                        <td width="20%" class="th1c_b">
                            會員編號
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_MAIL_SAME" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="15%" class="td2c_b">
                                    <!--點選-->
                                    <asp:RadioButton ID="RB_CHOOSE" runat="server" AutoPostBack="true" OnCheckedChanged="RB_CHOOSE_OnCheckedChanged" />
                                    <input type="hidden" id="HID_MB_MEMSEQ" runat="server" value='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                </td>
                                <td width="20%" class="td2c_b">
                                    <!--姓名-->
                                    <%#Container.DataItem("MB_NAME")%>
                                </td>
                                <td width="35%" class="td2_b">
                                    <!--通訊地址-->
                                    <%# Container.DataItem("MB_CITY")%><%# Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
                                </td>
                                <td width="20%" class="td2c_b">
                                    <!--會員編號-->
                                    <%#getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
                <div align="center" >
                    <asp:Button ID="btnModify" runat="server" CssClass="bt" Text="報名" />
                </div>
            </asp:PlaceHolder>

            <asp:DataGrid ID="dgRpt_Page2" Runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="0"
            BorderWidth="0" CellSpacing="0" CssClass="CRTable" ShowFooter="False" Visible="false">
                <Columns>
                    <asp:TemplateColumn>
                        <HeaderStyle CssClass="th1c_b"></HeaderStyle>
                        <HeaderTemplate>
                            點選
                        </HeaderTemplate>
                        <ItemStyle CssClass="td2c_b" HorizontalAlign="center"></ItemStyle>
                        <ItemTemplate>
                            <asp:RadioButton ID="rbData" runat="server" onclick="RadioButtonSelect(this);" />
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn>
                        <HeaderStyle CssClass="th1c_b"></HeaderStyle>
                        <HeaderTemplate>
                            法名/ 姓名
                        </HeaderTemplate>
                        <ItemStyle CssClass="td2c_b" HorizontalAlign="center"></ItemStyle>
                        <ItemTemplate>
                            <asp:Label ID="lbl_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn>
                        <HeaderStyle CssClass="th1c_b"></HeaderStyle>
                        <HeaderTemplate>
                            通訊地址
                        </HeaderTemplate>
                        <ItemStyle CssClass="td2c_b" HorizontalAlign="center"></ItemStyle>
                        <ItemTemplate>
                            <asp:Label ID="lbl_MB_CITY" runat="server" Text='<%#Container.DataItem("MB_CITY")%>'></asp:Label>
                            <asp:Label ID="lbl_MB_VLG" runat="server" Text='<%#Container.DataItem("MB_VLG")%>'></asp:Label>
                            <asp:Label ID="lbl_MB_ADDR" runat="server" Text='<%#Container.DataItem("MB_ADDR")%>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn>
                        <HeaderStyle CssClass="th1c_b"></HeaderStyle>
                        <HeaderTemplate>
                            會員編號
                        </HeaderTemplate>
                        <ItemStyle CssClass="td2c_b" HorizontalAlign="center"></ItemStyle>
                        <ItemTemplate>
                            <asp:Label ID="lbl_MB_AREA" runat="server" Text='<%#Container.DataItem("MB_AREA")%>'></asp:Label>
                            -
                            <asp:Label ID="lbl_MB_MEMSEQ" runat="server" Text='<%#Container.DataItem("MB_MEMSEQ")%>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
            </asp:DataGrid>
            <table width="100%" runat="server" id="tb_Page2_Button" visible="false">
                <tr>
                    <td width="100%" align="center">
                        <asp:Button ID="btn_Confirm" runat="server" Text="確定" CssClass="bt" />
                    </td>
                </tr>
            </table>

            <table class="CRTable_Top" width="100%" cellspacing="0" runat="server" id="tb_Page3" visible="false" >
                <tr>
                    <td class="th1_b" colspan="4">
                        會員編號 : 
                        <asp:Label ID="lbl_MB_AREA" runat="server" Text="" />-
                        <asp:Label ID="lbl_MB_MEMSEQ" runat="server" Text="" />
                        &nbsp;基本資料
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        課程名稱
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:Label ID="lbl_Class" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        法名/姓名
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_NAME" runat="server"></asp:TextBox>
                    </td>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        性別
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:RadioButtonList ID="rbtList_MB_SEX" runat="server" RepeatDirection="Horizontal">
                            <asp:ListItem Value="1" Text="男"></asp:ListItem>
                            <asp:ListItem Value="2" Text="女"></asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        出生年月日
                    </td>
                    <td id="TD_Y_1" runat="server" class="td2_b" width="35%">
                        西元
                        <asp:TextBox ID="txt_MB_BIRTH_YYY" runat="server" Width="10%"></asp:TextBox>年&nbsp;
                        <asp:TextBox ID="txt_MB_BIRTH_MM" runat="server" Width="8%"></asp:TextBox>月&nbsp;
                        <asp:TextBox ID="txt_MB_BIRTH_DD" runat="server" Width="8%"></asp:TextBox>日&nbsp;
                    </td>
                    <td id="TD_G_1_1" runat="server" class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        出家眾
                    </td>
                    <td id="TD_G_1_2" runat="server" class="td2_b" width="35%">
                        <asp:RadioButton ID="rbt_MB_MONK_Y" runat="server" Text="是" GroupName="MB_MONK" onclick="CheckMonk(this);"></asp:RadioButton>
                        <asp:RadioButton ID="rbt_MB_MONK_N" runat="server" Text="否" GroupName="MB_MONK" onclick="CheckMonk(this);"></asp:RadioButton>
                    </td>
                </tr>
                <tr runat="server" id="tr_MB_MONK_1" style="display:none">
                    <td class="th1c_b" width="15%">
                        法名
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_MONKNAME" runat="server" Width="30%"></asp:TextBox>
                    </td>
                    <td class="th1c_b" width="15%">
                        剃度/皈依恩師/戒師
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_MONKTECH" runat="server" Width="30%"></asp:TextBox>
                    </td>
                </tr>
                <tr runat="server" id="tr_MB_MONK_2" style="display:none">
                    <td class="th1c_b" width="15%">
                        傳承
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:DropDownList ID="dd_MB_EDUTYPE" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td class="th1c_b" width="15%">
                        常住/親近道場
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_MONKPLACE" runat="server" Width="30%"></asp:TextBox>
                    </td>
                </tr>
                <tr runat="server" id="tr_MB_MONK_3" style="display:none">
                    <td class="th1c_b" width="15%">
                        戒別
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:DropDownList ID="dd_MB_MONKTYPE" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td class="th1c_b" width="15%">
                        受戒日期
                    </td>
                    <td class="td2_b" width="35%">
                        西元
                        <asp:TextBox ID="txt_MB_MONKDATE_YYY" runat="server" Width="10%"></asp:TextBox>年&nbsp;
                        <asp:TextBox ID="txt_MB_MONKDATE_MM" runat="server" Width="8%"></asp:TextBox>月&nbsp;
                        <asp:TextBox ID="txt_MB_MONKDATE_DD" runat="server" Width="8%"></asp:TextBox>日
                    </td>
                </tr>
                <tr runat="server" id="tr_MB_MONK_4" style="display:none">
                    <td class="th1c_b" width="15%">
                        受戒地點
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:TextBox ID="txt_MB_MONKPLACE1" runat="server" Width="70%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        手機
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_MOBIL" runat="server" Width="40%"></asp:TextBox>
                    </td>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        電話
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_TEL_ZIP" runat="server" Width="10%" MaxLength="2"></asp:TextBox>
                        &nbsp;─&nbsp;
                        <asp:TextBox ID="txt_MB_TEL" runat="server" Width="40%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        E-mail
                    </td>
                    <td id="TD_Y_2" runat="server" class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_EMAIL" runat="server" Width="60%"></asp:TextBox>
                    </td>
                    <td id="TD_G_2_1" runat="server" class="th1c_b" width="15%">
                        <span style="color:Red">**</span>
                        身分證字號
                    </td>
                    <td id="TD_G_2_2" runat="server" class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_ID" runat="server" ></asp:TextBox>
                    </td>
                </tr>
                <tr >
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        參加動機或目的
                    </td>
                    <td class="td2_b" colspan="3">
                        <asp:TextBox ID="MB_OBJECT" runat="server" Columns="100" />
                    </td> 
                </tr>
                <tr >
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        是否曾參加過本中心課程?
                    </td>
                    <td class="td2_b" colspan="3">
                        <asp:RadioButtonList ID="JOINMBSC" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                            <asp:ListItem Text="是" Value="Y" />
                            <asp:ListItem Text="否" Value="N" />
                        </asp:RadioButtonList>
                    </td> 
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        學歷
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:DropDownList ID="dd_MB_EDU" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td class="th1c_b" width="15%">
                        介紹人/訊息來源
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_REFER" runat="server" ></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_3" runat="server" >
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        通訊地址
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:DropDownList ID="dd_MB_CITY" runat="server" AutoPostBack="true">
                        </asp:DropDownList>&nbsp;
                        <asp:DropDownList ID="dd_MB_VLG" runat="server">
                        </asp:DropDownList>&nbsp;
                        <asp:TextBox ID="txt_MB_ADDR" runat="server" Width="60%" ></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_4" runat="server">
                    <td class="th1c_b" width="15%">
                        戶籍地址
                        <asp:CheckBox ID="cb_Ditto" runat="server" Text="同上" AutoPostBack="True" />
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:DropDownList ID="dd_MB_CITY1" runat="server" AutoPostBack="true">
                        </asp:DropDownList>&nbsp;
                        <asp:DropDownList ID="dd_MB_VLG1" runat="server">
                        </asp:DropDownList>&nbsp;
                        <asp:TextBox ID="txt_MB_ADDR1" runat="server" Width="60%" ></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_5" runat="server">
                    <td class="th1c_b" width="15%">
                        語言
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:DataList ID="dl_MB_LANG" runat="server" RepeatDirection="Horizontal">
                            <ItemTemplate>
                                <asp:HiddenField ID="hid_MB_LANG" runat="server" Value='<%#Container.DataItem("VALUE")%>' />
                                <asp:CheckBox ID="cb_MB_LANG" runat="server" Text='<%#Container.DataItem("TEXT")%>'>
                                </asp:CheckBox>
                                <asp:TextBox ID="txt_MB_LANG" runat="server" Visible="false"></asp:TextBox>
                            </ItemTemplate>
                        </asp:DataList>
                    </td>
                </tr>
                <tr id="TR_G_6" runat="server">
                    <td class="th1c_b" width="15%">
                        專長
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:DropDownList ID="dd_MB_SPECIAL" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td class="th1c_b" width="15%">
                        職業
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:DropDownList ID="dd_MB_JOB" runat="server">
                        </asp:DropDownList>&nbsp;
                        職稱
                        <asp:DropDownList ID="dd_MB_JOBTITLE" runat="server">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">*</span>
                        宗教信仰
                    </td>
                    <td id="TD_Y_7" runat="server" class="td2_b" width="35%">
                        <asp:DropDownList ID="dd_MB_RELIGION" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td id="TD_G_7_1" runat="server" class="th1c_b" width="15%">
                        <span style="color:Red">**</span>
                        打鼾
                    </td>
                    <td id="TD_G_7_2" runat="server" class="td2_b" width="35%">
                        <asp:RadioButtonList ID="rbtList_MB_SNORE" runat="server">
                            <asp:ListItem Text="是" Value="1"></asp:ListItem>
                            <asp:ListItem Text="否" Value="2"></asp:ListItem>
                            <asp:ListItem Text="不知道" Value="3"></asp:ListItem>
                        </asp:RadioButtonList>
                        (為安排住宿)
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" width="15%">
                        公司/學校系級
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:TextBox ID="SCHOOL" runat="server" Columns="100" />
                    </td>
                </tr>
                <tr id="TR_G_8" runat="server">
                    <td class="th1c_b" width="15%">
                        <span style="color:Red">**</span>
                        身心狀況
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:DataList ID="dl_MB_SICK" runat="server">
                            <ItemTemplate>
                                <asp:HiddenField ID="hid_MB_SICK" runat="server" Value='<%#Container.DataItem("VALUE")%>' />
                                <asp:CheckBox ID="cb_MB_SICK" runat="server" Text='<%#Container.DataItem("TEXT")%>'>
                                </asp:CheckBox>
                                <asp:TextBox ID="txt_MB_SICK" runat="server" Visible="false"></asp:TextBox>
                            </ItemTemplate>
                        </asp:DataList>
                    </td>
                </tr>
                <tr id="TR_G_9" runat="server">
                    <td class="mtr" width="100%" colspan="4" align="center">
                        修持法門
                    </td>
                </tr>
                <tr id="TR_G_10" runat="server">
                    <td class="td2_b" width="100%" colspan="4">
                        您曾經修持過毗婆舍那禪法嗎?&nbsp;&nbsp;
                        <asp:RadioButton ID="rbt_MB_PIPOSHENA_Y" Text="是" runat="server" GroupName="MB_PIPOSHENA" />
                        <asp:RadioButton ID="rbt_MB_PIPOSHENA_N" Text="否" runat="server" GroupName="MB_PIPOSHENA" />
                        &nbsp;&nbsp;&nbsp;&nbsp;指導老師：
                        <asp:TextBox ID="txt_MB_TEACH" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_11" runat="server">
                    <td class="th1c_b" width="15%">
                        您目前的修持法門
                    </td>
                    <td class="td2_b" width="85%" colspan="3">
                        <asp:TextBox ID="txt_MB_FAMENNIAN" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_12" runat="server">
                    <td class="td2_b" width="100%" colspan="4">
                        您過去曾經參加過七天以上的禪修嗎?&nbsp;&nbsp;
                        <asp:RadioButton ID="rbt_MB_OVER7DAY_Y" Text="是" runat="server" GroupName="MB_OVER7DAY" />
                        <asp:RadioButton ID="rbt_MB_OVER7DAY_N" Text="否" runat="server" GroupName="MB_OVER7DAY" />
                        &nbsp;&nbsp;&nbsp;&nbsp;地點：
                        <asp:TextBox ID="txt_MB_PLACE" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_13" runat="server">
                    <td class="mtr" width="100%" colspan="4" align="center">
                        <span style="color:Red">**</span>
                        緊急聯絡人
                    </td>
                </tr>
                <tr id="TR_G_14" runat="server">
                    <td class="th1c_b" width="15%">
                        法名/姓名
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_EMGCONT" runat="server"></asp:TextBox>
                    </td>
                    <td class="th1c_b" width="15%">
                        電話/手機
                    </td>
                    <td class="td2_b" width="35%">
                        <asp:TextBox ID="txt_MB_CONTMOBIL" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr id="TR_G_15" runat="server" >
                    <td class="th1c_b" width="15%">
                        捐助物資/捐款
                    </td>
                    <td class="th1c_b" width="85%" colspan="3">
                        <asp:TextBox ID="MB_AMTMEMO" runat="server" />
                    </td>
                </tr>
            </table>
            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"--> 
            <table width="100%" runat="server" id="tb_Page3_Button" visible="false">
                <tr>
                    <td align="center">
                        <asp:Button ID="btn_Cancel" runat="server" Text="取消報名" CssClass="bt" Visible="false" />
                        &nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btn_Save" runat="server" Text="確定" CssClass="bt" />
                    </td>
                </tr>
            </table>
        
            <div style="color:red">
                說明
                <BR/>
                *:為所有課程必填項目,未填者將無法報名
                <BR/>
                **:七天以上(含)之禪修課程必填
            </div>

            <asp:HiddenField ID="hid_HaveMember" runat="server" Value="False" />
        </td>
    </tr>
</table>
    </form>
</body>
</html>
