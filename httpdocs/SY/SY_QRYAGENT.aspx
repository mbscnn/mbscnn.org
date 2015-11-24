<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="SY_QRYAGENT.aspx.vb" Inherits="MBSC.SY_QRYAGENT" %>

<%@ Register Src="../Module/UIControl/PagerMenu/PagerMenu.ascx" TagName="PagerMenu"
    TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>代理人查詢</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <script type="text/javascript">

        // 代理起訖時間檢核
        function CheckValidate() {
            var $startDate = $("#txtStartDate");
            var $startTime = $("#txtStartTime");
            var $endDate = $("#txtEndDate");
            var $endTime = $("#txtEndTime");
            var timeTest = new RegExp(/^((\d|([0,1]\d)|(2[0-3]))(:(0{1,2}|\d|([0-5][0-9])))?)$/);
            var dateTest = new RegExp(/^(\d{1,4}\.\d{1,2}\.\d{1,2})$/);
            var errorMsg = "";
            var $errorControl = null;
            var iCount = 0;

            if (jQuery.trim($startDate.val()) == "") {
                iCount++;
            } else {
                if (dateTest.test(jQuery.trim($startDate.val()))) {
                    var adDate = getADDate(jQuery.trim($startDate.val()));

                    if (!isValidDate(adDate.split(".")[0], adDate.split(".")[1], adDate.split(".")[2])) {
                        if ($errorControl == null) {
                            $errorControl = $startDate;
                        }

                        errorMsg += "代理起日期輸入格式有誤！\n";
                    }
                }
                else {
                    if ($errorControl == null) {
                        $errorControl = $startDate;
                    }

                    errorMsg += "代理起日期輸入格式有誤！\n";
                }
            }

            if (jQuery.trim($startTime.val()) == "") {
                iCount++;
            } else {
                if (!timeTest.test(jQuery.trim($startTime.val()))) {
                    if ($errorControl == null) {
                        $errorControl = $startTime;
                    }

                    errorMsg += "代理起時間格式輸入有誤！\n";
                }
            }

            if (jQuery.trim($endDate.val()) == "") {
                iCount++;
            } else {
                if (dateTest.test(jQuery.trim($endDate.val()))) {
                    var adDate = getADDate(jQuery.trim($endDate.val()));

                    if (!isValidDate(adDate.split(".")[0], adDate.split(".")[1], adDate.split(".")[2])) {
                        if ($errorControl == null) {
                            $errorControl = $endDate;
                        }

                        errorMsg += "代理訖日期輸入格式有誤！\n";
                    }
                }
                else {
                    if ($errorControl == null) {
                        $errorControl = $endDate;
                    }

                    errorMsg += "代理訖日期輸入格式有誤！\n";
                }
            }

            if (jQuery.trim($endTime.val()) == "") {
                iCount++;
            } else {
                if (!timeTest.test(jQuery.trim($endTime.val()))) {
                    if ($errorControl == null) {
                        $errorControl = $endTime;
                    }

                    errorMsg += "代理訖時間格式輸入有誤！\n";
                }
            }

            if (!(iCount == 0 || iCount == 4)) {
                alert("代理期間起訖必須同時輸入或不輸入！");

                $startDate[0].focus();

                return false;
            }

            if ($errorControl != null) {
                alert(errorMsg.substr(0, errorMsg.length - 2));

                if ($errorControl[0].type == "text") {
                    $errorControl.val("");
                }

                $errorControl[0].focus();
                return false;
            }

            return true;
        }

        // 由臺灣日期取得西元年
        function getADDate(twDate) {
            var year = twDate.split(".")[0];

            if (year.indexOf("-") > 0) {
                year = 1911 - parseInt(year.substr(year.indexOf("-") + 1));
            } else {
                year = 1911 + parseInt(year);
            }

            twDate = year.toString() + twDate.substr(twDate.indexOf("."));

            return twDate;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_QRYAGENT.docx')" />
    </div>
    <br />
    <div class="TaskTtitle">
        代理人查詢
    </div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
            <td>
                <div>
                    查詢種類：
                    <asp:RadioButton ID="rdoCurrent" AutoPostBack="true" OnCheckedChanged="rdoCurrent_CheckedChanged"
                        GroupName="agent" runat="server" />目前代理狀況
                    <asp:RadioButton ID="rdoHistory" AutoPostBack="true" OnCheckedChanged="rdoHistory_CheckedChanged"
                        GroupName="agent" runat="server" />歷史代理資料
                </div>
            </td>
        </tr>
    </table>
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
            <td>
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td class="td2" style="width: 15%;">
                            被代理人單位
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlStaffBranch" AutoPostBack="true" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="trStaff" runat="server" visible="false">
                        <td class="td2">
                            被代理人
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlStaff" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="trAgentStaffBranch" runat="server" visible="false">
                        <td class="td2">
                            代理人單位
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlAgentStaffBranch" AutoPostBack="true" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="trAgentStaff" runat="server" visible="false">
                        <td class="td2">
                            代理人
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlAgentStaff" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="trAgentDateTime" runat="server" visible="false">
                        <td class="td2">
                            代理期間
                        </td>
                        <td id="tdAgentDateTime">
                            <asp:TextBox ID="txtStartDate" runat="server"></asp:TextBox>
                            <img alt="" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"calendar.gif"%>"
                                style="cursor: hand" onclick="javascript:pedirFecha(document.all.item('txtStartDate'),'代理起日');" />
                            <asp:TextBox ID="txtStartTime" runat="server"></asp:TextBox>
                            ~
                            <asp:TextBox ID="txtEndDate" runat="server"></asp:TextBox>
                            <img alt="" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"calendar.gif"%>"
                                style="cursor: hand" onclick="javascript:pedirFecha(document.all.item('txtEndDate'),'代理訖日');" />
                            <asp:TextBox ID="txtEndTime" runat="server"></asp:TextBox>
                            <span style="color: Red;">(日期格式YYY.MM.DD)</span>
                        </td>
                    </tr>
                </table>
                <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
                    <tr>
                        <td align="center">
                            <asp:Button ID="btnConfirm" runat="server" Text="確定" OnClientClick="return CheckValidate();"
                                CssClass="bt" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <uc1:pagermenu id="PagerMenu1" runat="server" />
            </td>
        </tr>
        <tr>
            <td>
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td>
                            <asp:DataGrid ID="dgCurrentAgent" runat="server" CssClass="crtable" Width="100%"
                                BorderWidth="1px" CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False"
                                HeaderStyle-ForeColor="#000000" AllowSorting="True">
                                <Columns>
                                    <asp:TemplateColumn HeaderText="被代理人編號">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStaffId" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"STAFFID")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="代理人編號">
                                        <ItemTemplate>
                                            <asp:Label ID="lblAgentStaffId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"AGENTSTAFFID")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="12%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="代理單位">
                                        <ItemTemplate>
                                            <asp:Label ID="lblBrcName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRCNAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="代理起日">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStartDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STARTDATETIME","{0:yyyy.MM.dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="代理止日">
                                        <ItemTemplate>
                                            <asp:Label ID="lblEndDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ENDDATETIME","{0:yyyy.MM.dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="設定人">
                                        <ItemTemplate>
                                            <asp:Label ID="lblCreateUser" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CREATEUSER")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="設定日">
                                        <ItemTemplate>
                                            <asp:Label ID="lblCreateDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CREATETIME","{0:yyyy.MM.dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="取消日">
                                        <ItemTemplate>
                                            <asp:Label ID="lblCancelDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CANCELTIME","{0:yyyy.MM.dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    <asp:HiddenField ID="hidScript" runat="server" />
    </form>
</body>
</html>
