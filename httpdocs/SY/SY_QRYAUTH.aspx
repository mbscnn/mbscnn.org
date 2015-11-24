<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_QRYAUTH.aspx.vb" Inherits="MBSC.SY_QRYAUTH" %>

<%@ Register Src="../Module/UIControl/PagerMenu/PagerMenu.ascx" TagName="PagerMenu"
    TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>人員分派查詢</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <script type="text/javascript">
        function CheckValidate() {
            var $startYear = $("#ddlStartYear");
            var $startMonth = $("#ddlStartMonth");
            var $startDay = $("#ddlStartDay");
            var $endYear = $("#ddlEndYear");
            var $endMonth = $("#ddlEndMonth");
            var $endDay = $("#ddlEndDay");
            var erroeMsg = "";
            var $errorControl = null;

            if ($startYear.val() == "0") {
                $errorControl = $startYear;

                erroeMsg = "請選擇開始年份！\n";
            }

            if ($startMonth.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $startMonth;
                }

                erroeMsg += "請選擇開始月份！\n";
            }

            if ($startDay.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $startDay;
                }

                erroeMsg += "請選擇開始日！\n";
            }


            if ($endYear.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $endYear;
                }

                erroeMsg += "請選擇截止年份！\n";
            }

            if ($endMonth.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $endMonth;
                }

                erroeMsg += "請選擇截止月份！\n";
            }

            if ($endDay.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $endDay;
                }

                erroeMsg += "請選擇截止日！\n";
            }


            if ($startYear.val() != "0" && $startMonth.val() != "0" && $startDay.val() != "0" && $endYear.val() != "0" && $endMonth.val() != "0" && $endDay.val() != "0") {
                if (($startYear.val() + $startMonth.val() + $startDay.val()) > ($endYear.val() + $endMonth.val() + $endDay.val())) {
                    if ($errorControl == null) {
                        $errorControl = $startYear;
                    }

                    erroeMsg += "起期間不能大於訖期間！\n";
                }
            }

            if ($errorControl != null) {
                alert(erroeMsg);

                $errorControl[0].focus();

                return false;
            }

            return true;
        }

        function pop_DetailPage(sURL) {
            var strFeatures = "dialogWidth:640px;dialogHeight:480px;center=yes;help=no;status=no;resizable=no";
            var st = new Array();
            window.open(sURL, st, strFeatures);
            //location.reload();
        };
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_QRYAUTH.docx')" />
    </div>
    <br />
    <div class="TaskTtitle">
        人員分派查詢
    </div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
            <td colspan="2">
                單位：
                <asp:DropDownList ID="ddlBranch" AutoPostBack="true" runat="server">
                </asp:DropDownList>
        </tr>
        <tr>
            <td colspan="2">
                查詢種類：
                <asp:RadioButton ID="rdoUser" AutoPostBack="true" Checked="true" GroupName="qryType"
                    runat="server" />依人員
                <asp:RadioButton ID="rdoRole" AutoPostBack="true" GroupName="qryType" runat="server" />依角色
                <asp:RadioButton ID="rdoChange" AutoPostBack="true" GroupName="qryType" runat="server" />異動查詢
            </td>
        </tr>
        <tr id="trTime" visible="false" runat="server">
            <td>
                期間：自
                <asp:DropDownList ID="ddlStartYear" runat="server">
                </asp:DropDownList>
                年
                <asp:DropDownList ID="ddlStartMonth" runat="server">
                </asp:DropDownList>
                月<asp:DropDownList ID="ddlStartDay" runat="server">
                </asp:DropDownList>
                日 ～
                <asp:DropDownList ID="ddlEndYear" runat="server">
                </asp:DropDownList>
                年
                <asp:DropDownList ID="ddlEndMonth" runat="server">
                </asp:DropDownList>
                月<asp:DropDownList ID="ddlEndDay" runat="server">
                </asp:DropDownList>
                日
                <asp:Button ID="btSearch" runat="server" Text="查詢" OnClientClick="return CheckValidate();"
                    class="bt" />
            </td>
        </tr>
        <tr id="trPagerMenu" runat="server">
            <td>
                <uc1:pagermenu id="PagerMenu1" runat="server" />
            </td>
        </tr>
        <tr id="trUser" runat="server">
            <td colspan="2">
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td>
                            <asp:DataGrid ID="dgUser" runat="server" CssClass="crtable" Width="100%" BorderWidth="1px"
                                CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                                AllowSorting="True">
                                <Columns>
                                    <asp:TemplateColumn HeaderText="員工姓名">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStaffName" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"STAFF")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="20%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="20%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="角色名稱">
                                        <ItemTemplate>
                                            <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="80%" CssClass="td2_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="80%" />
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr id="trRole" visible="false" runat="server">
            <td colspan="2">
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td>
                            <asp:DataGrid ID="dgRole" runat="server" CssClass="crtable" Width="100%" BorderWidth="1px"
                                CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                                AllowSorting="True">
                                <Columns>
                                    <asp:TemplateColumn HeaderText="角色名稱">
                                        <ItemTemplate>
                                            <asp:Label ID="lblRoleName" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"ROLENAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="20%" CssClass="td2_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="20%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="員工姓名">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStaffName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STAFF")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="80%" CssClass="td2_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="80%" />
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr id="trChange" visible="false" runat="server">
            <td colspan="2">
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td>
                            <asp:DataGrid ID="dgChange" runat="server" CssClass="crtable" Width="100%" BorderWidth="1px"
                                CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                                AllowSorting="True">
                                <Columns>
                                    <asp:TemplateColumn HeaderText="行員姓名">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStaffName" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"STAFF")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="17%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="17%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="角色名稱">
                                        <ItemTemplate>
                                            <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="17%" CssClass="td2_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="17%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="異動">
                                        <ItemTemplate>
                                            <asp:Label ID="lblOperation" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"OPERATION")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="15%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="15%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="異動日期">
                                        <ItemTemplate>
                                            <asp:Label ID="lblOperateDate" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ENDTIME","{0:yyyy/MM/dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="17%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="17%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="維護主管一">
                                        <ItemTemplate>
                                            <asp:Label ID="lblEndDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STAFFONE")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="17%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="17%" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="維護主管二">
                                        <ItemTemplate>
                                            <asp:Label ID="lblCreateUser" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STAFFTWO")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="17%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" Width="17%" />
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>                    
                </table>
            </td>
        </tr>         
        <tr id="trExcel" visible="false" runat="server">
            <td colspan="2" align="center" >
                <asp:Button ID="btntoexcel" runat="server" Text="另存EXCEL" CssClass="bt" />
            </td>
        </tr>    
    </table>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
