<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDAGENT.aspx.vb" Inherits="MBSC.SY_UPDAGENT" %>

<%@ Register src="../Module/UIControl/PagerMenu/PagerMenu.ascx" tagname="PagerMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�N�z�H�]�w</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--�@�Ϊ�SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
	<!-- #include virtual="~/inc/MBSCJS.inc" --> 

    <script type="text/javascript">

        // �N�z�_�W�ɶ��ˮ�
        function CheckValidate() {
            var $staff = $("#ddlStaff");
            var $agentStaff = $("#ddlAgentStaff");
            var $staffBranch = $("#ddlStaffBranch");
            var $agentStaffBranch = $("#ddlAgentStaffBranch");
            var $startDate = $("#txtStartDate");
            var $startTime = $("#txtStartTime");
            var $endDate = $("#txtEndDate");
            var $endTime = $("#txtEndTime");
            var timeTest = new RegExp(/^((\d|([0,1]\d)|(2[0-3]))(:(0{1,2}|\d|([0-5][0-9])))?)$/);
            var dateTest = new RegExp(/^(\d{1,4}\.\d{1,2}\.\d{1,2})$/);
            var errorMsg = "";
            var $errorControl = null;

            if ($staffBranch.val() == "0") {
                $errorControl = $staffBranch;
                errorMsg += "�п�ܳQ�N�z�H���I\n";
            }

            if ($staff.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $staff;
                }
                
                errorMsg += "�п�ܳQ�N�z�H�I\n";
            }

            if ($agentStaffBranch.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $agentStaffBranch;
                }

                errorMsg += "�п�ܥN�z�H���I\n";
            }

            if ($agentStaff.val() == "0") {
                if ($errorControl == null) {
                    $errorControl = $agentStaff;
                }

                errorMsg += "�п�ܥN�z�H�I\n";
            }

            if ($staff.val().split(";")[0] == $agentStaff.val().split(";")[0] && $agentStaff.val() != "0") {
                if ($errorControl == null) {
                    $errorControl = $agentStaff;
                }

                errorMsg += "�Q�N�z�H�P�N�z�H���i��ܦP�@�H�I\n";
            }

            if (jQuery.trim($startDate.val()) == "") {
                if ($errorControl == null) {
                    $errorControl = $startDate;
                }

                errorMsg += "�п�J�N�z�_����I\n";
            } else {
                if (dateTest.test(jQuery.trim($startDate.val()))) {
                    var adDate = getADDate(jQuery.trim($startDate.val()));

                    if (!isValidDate(adDate.split(".")[0], adDate.split(".")[1], adDate.split(".")[2])) {
                        if ($errorControl == null) {
                            $errorControl = $startDate;
                        }

                        errorMsg += "�N�z�_�����J�榡���~�I\n";
                    }
                }
                else {
                    if ($errorControl == null) {
                        $errorControl = $startDate;
                    }

                    errorMsg += "�N�z�_�����J�榡���~�I\n";
                }
            }

            if (jQuery.trim($startTime.val()) == "") {
                if ($errorControl == null) {
                    $errorControl = $startTime;
                }

                errorMsg += "�п�J�N�z�_�ɶ��I\n";
            } else {
                if (!timeTest.test(jQuery.trim($startTime.val()))) {
                    if ($errorControl == null) {
                        $errorControl = $startTime;
                    }

                    errorMsg += "�N�z�_�ɶ��榡��J���~�I\n";
                }
            }

            if (jQuery.trim($endDate.val()) == "") {
                if ($errorControl == null) {
                    $errorControl = $endDate;
                }

                errorMsg += "�п�J�N�z�W����I\n";
            }
            else {
                if (dateTest.test(jQuery.trim($endDate.val()))) {
                    var adDate = getADDate(jQuery.trim($endDate.val()));

                    if (!isValidDate(adDate.split(".")[0], adDate.split(".")[1], adDate.split(".")[2])) {
                        if ($errorControl == null) {
                            $errorControl = $endDate;
                        }

                        errorMsg += "�N�z�W�����J�榡���~�I\n";
                    }
                }
                else {
                    if ($errorControl == null) {
                        $errorControl = $endDate;
                    }

                    errorMsg += "�N�z�W�����J�榡���~�I\n";
                }
            }

            if (jQuery.trim($endTime.val()) == "") {
                if ($errorControl == null) {
                    $errorControl = $endTime;
                }

                errorMsg += "�п�J�N�z�W�ɶ��I\n";
            } else {
                if (!timeTest.test(jQuery.trim($endTime.val()))) {
                    if ($errorControl == null) {
                        $errorControl = $endTime;
                    }

                    errorMsg += "�N�z�W�ɶ��榡��J���~�I\n";
                }
            }

            if ($errorControl != null) {
                alert(errorMsg);

                if ($errorControl[0].type == "text") {
                    $errorControl.val("");
                }

                $errorControl[0].focus();
                return false;
            }

            return true;
        }

        // �ѻO�W������o�褸�~
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
            onclick="javascript:window.open('./help/SY_UPDAGENT.docx')" />
    </div>
    <br />
    <div class ="TaskTtitle">
     �N�z�H�]�w
    </div>
    <br />
    <!--��ܤW��~�ؽu-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr id="trAgentSet" runat="server" visible="false">
            <td>
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td class="td2">
                            �Q�N�z�H���
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlStaffBranch" AutoPostBack="true" runat="server" Width ="210px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="td2">
                            �Q�N�z�H
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlStaff" runat="server" Width ="140px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="td2">
                            �N�z�H���
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlAgentStaffBranch" AutoPostBack="true" runat="server" Width ="210px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="td2">
                            �N�z�H
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlAgentStaff" runat="server" Width ="140px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="td2">
                            �N�z����
                        </td>
                        <td>
                            <asp:TextBox ID="txtStartDate" runat="server"></asp:TextBox>
                            <img alt="" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"calendar.gif"%>"
                                style="cursor: hand" onclick="javascript:pedirFecha(document.all.item('txtStartDate'),'�N�z�_��');" />
                            <asp:TextBox ID="txtStartTime" runat="server"></asp:TextBox>
                            ~
                            <asp:TextBox ID="txtEndDate" runat="server"></asp:TextBox>
                            <img alt="" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"calendar.gif"%>"
                                style="cursor: hand" onclick="javascript:pedirFecha(document.all.item('txtEndDate'),'�N�z�W��');" />
                            <asp:TextBox ID="txtEndTime" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                </table>
                <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
                    <tr>
                        <td align="center">
                            <asp:Button ID="btnSave" runat="server" Text="�x�s" OnClientClick="return CheckValidate();"
                                CssClass="bt" />
                            <asp:Button ID="btnCancel" runat="server" Text="����" CssClass="bt" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr id="trAgentSetDetail" runat="server">
            <td>
                <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
                    <tr>
                        <td align="right">
                            <asp:Button ID="btnAdd" runat="server" Text="�s�W" CssClass="bt" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <uc1:PagerMenu ID="PagerMenu1" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DataGrid ID="dgAgentInfo" runat="server" CssClass="crtable" Width="100%" BorderWidth="1px"
                                CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                                AllowSorting="True">
                                <Columns>
                                    <asp:TemplateColumn HeaderText="�Q�N�z�H�s��">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStaffId" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"STAFFID")%>'></asp:Label>
                                            <asp:HiddenField ID="hidBraDepNo" runat="server" Value='<%# DataBinder.Eval(Container.DataItem,"BRADEPNO")%>' />
                                            <asp:HiddenField ID="hidAgentBraDepNo" runat="server" Value='<%# DataBinder.Eval(Container.DataItem,"AGENTBRADEPNO")%>' />
                                            <asp:HiddenField ID="hidCreateTime" runat="server" Value='<%# DataBinder.Eval(Container.DataItem,"CREATETIME","{0:yyyy.MM.dd HH:mm:ss.fff}")%>' />
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="�Q�N�z�H�m�W">
                                        <ItemTemplate>
                                            <asp:Label ID="lblUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="�N�z�H�s��">
                                        <ItemTemplate>
                                            <asp:Label ID="lblAgentStaffId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"AGENTSTAFFID")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="12%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="�N�z�H�m�W">
                                        <ItemTemplate>
                                            <asp:Label ID="lblAgentUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"AGENTUSERNAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="�N�z���">
                                        <ItemTemplate>
                                            <asp:Label ID="lblBrcName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRCNAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="�N�z�_��">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStartDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STARTDATETIME","{0:yyyy.MM.dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="�N�z���">
                                        <ItemTemplate>
                                            <asp:Label ID="lblEndDateTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ENDDATETIME","{0:yyyy.MM.dd HH:mm}")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="13%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="����">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgbtnCancel" runat="server" CommandName="Cancel" ImageUrl="~/img/del.gif" />
                                        </ItemTemplate>
                                        <ItemStyle Width="9%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="th1c_b" />
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <!--��ܤU��~�ؽu-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
