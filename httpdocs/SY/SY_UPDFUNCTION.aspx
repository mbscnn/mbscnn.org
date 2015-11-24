<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDFUNCTION.aspx.vb"
    Inherits="MBSC.SY_UPDFUNCTION" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>交易建立</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <base target="_self" />
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_UPDFUNCTION.docx')" />
    </div>
    <br />
    <div id="divBranch" runat="server" visible="false">
        <table id="Table3" width="100%" runat="server">
            <tr align="right">
                <td colspan="2">
                    案件編號:<asp:Label ID="lblCaseId" runat="server"></asp:Label>
                </td>
            </tr>
            <tr align="right">
                <td colspan="2">
                    申請單位:
                    <asp:Label ID="lblBranch" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    <div class="TaskTtitle">
        交易建立
    </div>
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <div id="divNormal" runat="server">
        <table width="100%">
            <tr>
                <td align="left" width="30%" valign="top">
                    <asp:TreeView ID="treeView" runat="server" TabIndex="5">
                        <RootNodeStyle ImageUrl="~/img/folder.png" />
                        <ParentNodeStyle ImageUrl="~/img/folder.png" />
                        <LeafNodeStyle ImageUrl="~/img/file.gif" />
                    </asp:TreeView>
                </td>
                <td align="right" valign="top">
                    <asp:Button ID="btnConfirmSend" runat="server" Text="確認送出" CssClass="bt" TabIndex="10" />&nbsp;&nbsp;
                    <asp:Button ID="btnCancelAll" runat="server" Text="全部取消" CssClass="bt" TabIndex="15" />
                    <table id="Table1" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                        <tr>
                            <td>
                                上層項目
                            </td>
                            <td>
                                <asp:Label ID="lblParentName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr runat="server" id="trId" visible="false">
                            <td style="width: 15%">
                                編號
                            </td>
                            <td style="width: 85%">
                                <asp:Label ID="lblID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr runat="server" id="trName">
                            <td>
                                名稱
                            </td>
                            <td>
                                <asp:Label ID="lblName" runat="server"></asp:Label>
                                <asp:TextBox ID="txtName" runat="server" MaxLength="100" Visible="false" TabIndex="16"
                                    AutoPostBack="true"></asp:TextBox>
                            </td>
                        </tr>
                        <tr runat="server" id="trFuncUrl" visible="false">
                            <td>
                                路徑
                            </td>
                            <td>
                                <asp:TextBox ID="txtFuncUrl" runat="server" MaxLength="100" TabIndex="17" AutoPostBack="True"
                                    Width="300px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr runat="server" id="trSortCtrl" visible="false">
                            <td>
                                排序控制
                            </td>
                            <td>
                                <asp:TextBox ID="txtSortCtrl" runat="server" TabIndex="18" AutoPostBack="true" Width="50px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trHoFlag" runat="server" visible="false">
                            <td>
                                屬性
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoListHoFlag" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="19" AutoPostBack="true">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trDisable" visible="false" runat="server">
                            <td>
                                啟用狀態
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoDisable" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="20" AutoPostBack="true">
                                    <asp:ListItem Selected="True" Value="0">啟用</asp:ListItem>
                                    <asp:ListItem Value="1">停用</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="right">
                                <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="bt" TabIndex="21" />&nbsp;&nbsp;
                                <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="bt" TabIndex="22" Visible="false" />&nbsp;&nbsp;
                                <asp:Button ID="btnDelete" runat="server" Text="刪除" CssClass="bt" TabIndex="23" Visible="false" />&nbsp;&nbsp;
                                <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="bt" TabIndex="24" Visible="false" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <!--記錄當前資料的父節點的ID-->
        <asp:HiddenField ID="hidParentId" runat="server" />
        <!--記錄當前資料的父節點的Name-->
        <asp:HiddenField ID="hidParentName" runat="server" />
        <!--記錄當前資料的ID-->
        <asp:HiddenField ID="hidId" runat="server" />
        <!--記錄當前資料的Name-->
        <asp:HiddenField ID="hidName" runat="server" />
        <!--記錄是否有點擊新增按鈕-->
        <asp:HiddenField ID="hidFlag" runat="server" />
        <!--記錄父系統-->
        <asp:HiddenField ID="hidSysId" runat="server" />
        <!--記錄子系統-->
        <asp:HiddenField ID="hidSubSysId" runat="server" />
        <asp:HiddenField ID="hidNodeParent" runat="server" />
        <asp:HiddenField ID="hidParent" runat="server" />
        <asp:HiddenField ID="hidOldXmlData" runat="server" />
        <asp:HiddenField ID="hidNewXmlData" runat="server" />
        <asp:HiddenField ID="hidOldData" runat="server" />
        <asp:HiddenField ID="hidValue" runat="server" />
        <asp:HiddenField ID="hidParentFuncList" runat="server" />
          <asp:HiddenField ID="hidParId" runat="server" />
    </div>
    <div id="divCheck" runat="server" visible="false">
        <table width="100%" runat="server" id="tdCheck">
            <tr>
                <td align="left" width="20%" valign="top">
                    <asp:TreeView ID="treeViewCheck" runat="server" TabIndex="15">
                    </asp:TreeView>
                </td>
                <td align="right" valign="top" style="width: 80%">
                    <table id="Table2" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                        <tr>
                            <td colspan="2">
                                修改前
                            </td>
                            <td colspan="2">
                                修改后
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%">
                                上層項目
                            </td>
                            <td>
                                <asp:Label ID="lblParentNameBefore" runat="server" TabIndex="20"></asp:Label>
                            </td>
                            <td>
                                上層項目
                            </td>
                            <td>
                                <asp:Label ID="lblParentNameAfter" runat="server" TabIndex="25"></asp:Label>
                            </td>
                        </tr>
                        <tr id="Tr1" runat="server">
                            <td>
                                編號
                            </td>
                            <td>
                                <asp:Label ID="lblIDBefore" runat="server" TabIndex="30"></asp:Label>
                            </td>
                            <td>
                                編號
                            </td>
                            <td>
                                <asp:Label ID="lblIDAfter" runat="server" TabIndex="35"></asp:Label>
                            </td>
                        </tr>
                        <tr id="Tr2" runat="server">
                            <td>
                                名稱
                            </td>
                            <td>
                                <asp:Label ID="lblNameBefore" runat="server" TabIndex="40"></asp:Label>
                                <asp:TextBox ID="txtNameBefore" runat="server" MaxLength="100" Visible="false" TabIndex="45"></asp:TextBox>
                            </td>
                            <td>
                                名稱
                            </td>
                            <td>
                                <asp:Label ID="lblNameAfter" runat="server" TabIndex="50"></asp:Label>
                                <asp:TextBox ID="txtNameAfter" runat="server" MaxLength="100" Visible="false" TabIndex="55"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="tr3" runat="server">
                            <td>
                                路徑
                            </td>
                            <td>
                                <asp:TextBox ID="txtFuncUrlBefore" runat="server" MaxLength="100" TabIndex="60" Enabled="false"></asp:TextBox>
                            </td>
                            <td>
                                路徑
                            </td>
                            <td>
                                <asp:TextBox ID="txtFuncUrlAfter" runat="server" MaxLength="100" TabIndex="65" Enabled="false"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="tr4" runat="server">
                            <td>
                                排序控制
                            </td>
                            <td>
                                <asp:TextBox ID="txtSortCtrlBefore" runat="server" TabIndex="70" Enabled="false"></asp:TextBox>
                            </td>
                            <td>
                                排序控制
                            </td>
                            <td>
                                <asp:TextBox ID="txtSortCtrlAfter" runat="server" TabIndex="75" Enabled="false"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="tr5" runat="server">
                            <td>
                                屬性
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoListHoFlagBefore" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="80" Enabled="false" Visible="false">
                                </asp:RadioButtonList>
                            </td>
                            <td>
                                屬性
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoListHoFlagAfter" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="85" Enabled="false" Visible="false">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr6" runat="server">
                            <td>
                                啟用狀態
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoDisableBefore" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="90" Enabled="false">
                                    <asp:ListItem Selected="True" Value="0">啟用</asp:ListItem>
                                    <asp:ListItem Value="1">停用</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td>
                                啟用狀態
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoDisableAfter" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="95" Enabled="false">
                                    <asp:ListItem Selected="True" Value="0">啟用</asp:ListItem>
                                    <asp:ListItem Value="1">停用</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" align="center">
                                <asp:Button ID="btnArgee" runat="server" Text="同意" CssClass="bt" TabIndex="100" />&nbsp;&nbsp;
                                <asp:Button ID="btnNotArgee" runat="server" Text="不同意" CssClass="bt" TabIndex="105" />&nbsp;&nbsp;
                                <asp:Button ID="btnReviseFlow" runat="server" Text="修正補充" CssClass="bt" TabIndex="110" />&nbsp;&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
