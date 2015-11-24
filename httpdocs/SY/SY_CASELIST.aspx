<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_CASELIST.aspx.vb" Inherits="MBSC.SY_CASELIST" %>

<%@ Register Src="../Module/UIControl/PagerMenu/PagerMenu.ascx" TagName="PagerMenu"
    TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>待處理工作</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY JS-->.inc
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <script type="text/javascript">
        // 只能選擇一筆案件資料
        function ChangeCheck(obj) {
            var $dgCase = $("#dgCase");
            var $rdoCheck = $dgCase.find("input:radio");

            if (obj.checked) {
                $rdoCheck.removeAttr("checked");
                obj.checked = true;
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_CASELIST.docx')" />
    </div>
    <div class="TaskTtitle">
        待處理工作
    </div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoantableStart.inc"-->
    <table width="100%" border="1px" cellspacing="0" cellpadding="0" class="mtr">
        <tr>
            <td align="right" class="th1">
                日期：<asp:Label ID="lblDate" runat="server" />
            </td>
        </tr>
        <tr>
            <td>
                <uc1:pagermenu id="PagerMenu1" runat="server" />
            </td>
        </tr>
        <tr>
            <td>
                <asp:DataGrid ID="dgCase" runat="server" CssClass="CRTABLE" Width="100%" BorderWidth="1px" OnSortCommand="sortCommand"
                    CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000" 
                    AllowSorting="True">
                    <AlternatingItemStyle CssClass="DataGrid_AlternatingItemStyle"></AlternatingItemStyle>
                    <HeaderStyle ForeColor="Black"></HeaderStyle>
                    <Columns>
                        <asp:TemplateColumn HeaderText="">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="4%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:RadioButton ID="rdoChk" runat="server" onclick="ChangeCheck(this);" />
                                <asp:HiddenField ID="hidFormUrl" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"FORMURL")%>' />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="案件編號" SortExpression="CASEID">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkbtnCaseId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CASEID")%>'></asp:LinkButton>
                                <asp:HiddenField ID="hidSubFlowSeq" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"SUBFLOW_SEQ")%>' />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="核准編號" SortExpression="APV_CAS_ID">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblApvCasId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"APV_CAS_ID")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="業務類別" SortExpression="FLOW_CNAME">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblFlowCName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"FLOW_CNAME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="授信戶名稱" SortExpression="CPL_APL_ID">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblCplAplId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CPL_APL_ID")%>'>
                                </asp:Label>
                                <asp:Label ID="lblCplAplNam" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CPL_APL_NAM")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="案件狀態" SortExpression="STEP_NAME">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblStepName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STEP_NAME")%>'>
                                </asp:Label>
                                <asp:HiddenField ID="hidStepNo" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"STEP_NO")%>' />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="申請時間" SortExpression="CASESTTIME">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblCaseStTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CASESTTIME","{0:yyyy/MM/dd (HH:mm)}")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="收件時間" SortExpression="STARTTIME">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblStartTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STARTTIME","{0:yyyy/MM/dd (HH:mm)}")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="送件步驟" SortExpression="PR_STEPNAME">
                             <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>                  
                            <ItemTemplate>
                                <asp:Label ID="LblPreviousStep" runat="server" 
                                    Text='<%# DataBinder.Eval(Container.DataItem,"PR_STEPNAME") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="送件單位" SortExpression="PR_BRCNAME">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblPreBrcName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"PR_BRCNAME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>

                        <asp:TemplateColumn HeaderText="申請單位" SortExpression="BRCNAME">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" 
                                Width="10%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblBrcName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRCNAME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>
            </td>
        </tr>
    </table>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoantableEnd.inc"-->
    </form>
</body>
</html>
