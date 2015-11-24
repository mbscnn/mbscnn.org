<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDBRANCH.aspx.vb"
    Inherits="MBSC.SY_UPDBRANCH" %>

<%@ Register Src="../Module/UIControl/Addr/ZipCode.ascx" TagName="ZipCode" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>組織維護作業</title>
    <base target="_self" />
    <!--共用的SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <script type="text/javascript">
        // 輸入驗證
        function CheckValidate() {
            var branchName = document.getElementById("txtBranchName");
            var braTel = document.getElementById("txtBraTel");
            var brTel = document.getElementById("txtBrTel");
            var brid = document.getElementById("txtBrid");
            var mgrBrids = document.getElementById("txtMgrBrid");
            var braTelTest = new RegExp(/^\d{0,4}$/);
            var brTelTest = new RegExp(/^\d{0,10}$/);
            var errorMsg = "";
            var errorControl = null;

            // ?位名?
            if (branchName != null) {
                //  單位名稱必填驗證
                if (TrimText(branchName.value) == "") {
                    errorControl = branchName;

                    errorMsg = "請輸入單位名稱！\n";
                }
            }

            // ?位代?
            if (TrimText(brid.value) == "") {
                if (errorControl == null) {
                    errorControl = brid;
                }

                errorMsg += "請輸入單位代碼！\n";
            }

            // 管理?位
            if (TrimText(mgrBrids.value) != "") {
                var mgrBrid = mgrBrids.value.replace("，", ",").split(",");

                for (var i = 0; i < mgrBrid.length; i++) {
                    if (!isNumber(mgrBrid[i])) {
                        if (errorControl == null) {
                            errorControl = mgrBrids;
                        }

                        errorMsg += "管理單位只能輸入數字（多個時用‘，’隔?）！\n";
                        
                        break;
                    }
                }
            }

            // 電話區號輸入驗證
            if (!braTelTest.test(TrimText(braTel.value))) {
                if (errorControl == null) {
                    errorControl = braTel;
                }

                errorMsg += "電話區最多只能輸入4碼整數！\n";
            }

            // 電話驗證 
            if (!brTelTest.test(TrimText(brTel.value))) {
                if (errorControl == null) {
                    errorControl = brTel;
                }

                errorMsg += "電話區最多只能輸入10碼整數！\n";
            }

            if (errorControl != null) {
                alert(errorMsg);
                errorControl.focus();
                errorControl.value = "";
                return false;
            }

            return true;
        }

        // 去空格
        function TrimText(text) {
            return (text || "").replace(RegExp(/^\s+|\s+$/g), "");
        }

        // 檢查是否為數字
        function isNumber(strNum) {
            var num = new Number(stripCommas(strNum));
            if (isNaN(num)) {
                return false;
            }
            return true;
        }

        // ?“，”替??空
        function stripCommas(numString) {
            var re = /,/g;
            if (typeof numString == "number") {
                return numString;
            }
            return numString.replace(re, "");
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_UPDBRANCH.docx')" />
    </div>
    <br />
    <div class="TaskTtitle">
        組織維護作業
    </div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table width="100%">
        <tr>
            <td colspan="2">
                <asp:RadioButton ID="rdoAll" AutoPostBack="true" runat="server" GroupName="branch"
                    TabIndex="5" />
                全部
                <asp:RadioButton ID="rdoBranch" AutoPostBack="true" Checked="true" runat="server"
                    GroupName="branch" TabIndex="10" />
                依單位
                <asp:DropDownList ID="ddlBranch" AutoPostBack="true" runat="server" TabIndex="15">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td>
                <table id="tblContent" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                    <tr>
                        <td align="left" width="30%" valign="top">
                            <asp:TreeView ID="tvBranch" runat="server" ExpandDepth ="2">
                                <RootNodeStyle ImageUrl="~/img/company1.png" />
                                <ParentNodeStyle ImageUrl="~/img/bank2.png" />
                                <LeafNodeStyle ImageUrl="~/img/company3.png" />
                            </asp:TreeView>
                        </td>
                        <td align="right" valign="top">
                            <table id="tblDetail" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                                <tr>
                                    <td style="width: 30%">
                                        上層單位
                                    </td>
                                    <td>
                                        <asp:Label ID="lblParentName" runat="server"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        單位名稱
                                    </td>
                                    <td>
                                        <asp:Label ID="lblBranchName" runat="server"></asp:Label>
                                        <asp:TextBox ID="txtBranchName" runat="server" MaxLength="30" Visible="false" 
                                            TabIndex="20" Width="214px"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="trBrid" runat="server">
                                    <td>
                                        單位代碼
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtBrid" runat="server" MaxLength = "3" Width ="50px" TabIndex="25"></asp:TextBox>
                                        <asp:HiddenField ID="hidBraDepNo" runat="server" />
                                        <asp:HiddenField ID="hidBrid" runat="server" />
                                    </td>
                                </tr>
                                <tr id="trMgrBrid" runat="server">
                                    <td>
                                        管理單位
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtMgrBrid" runat="server" Width ="50px" TabIndex="30"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="trAddr" runat="server">
                                    <td>
                                        地址
                                    </td>
                                    <td>
                                        <uc1:zipcode id="brAddr" runat="server"/>
                                    </td>
                                </tr>
                                <tr id="trTel" runat="server">
                                    <td>
                                        電話
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtBraTel" runat="server" MaxLength="4" Width="10%" TabIndex="40"></asp:TextBox>
                                        －
                                        <asp:TextBox ID="txtBrTel" runat="server" MaxLength="10" TabIndex="45"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="trProperty" runat="server">
                                    <td>
                                        屬性
                                    </td>
                                    <td>
                                        <asp:RadioButtonList ID="rdolProperty" runat="server" RepeatDirection="Horizontal" TabIndex="50">
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr id="trStatus" runat="server">
                                    <td>
                                        啟用狀態
                                    </td>
                                    <td>
                                        <asp:RadioButtonList ID="rdoDisable" runat="server" RepeatDirection="Horizontal"
                                            TabIndex="55">
                                            <asp:ListItem Selected="True" Value="0">啟用</asp:ListItem>
                                            <asp:ListItem Value="1">停用</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" align="right">
                                        <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="bt" TabIndex="60" />&nbsp;&nbsp;
                                        <asp:Button ID="btnDelete" runat="server" Text="刪除" CssClass="bt" OnClientClick="return confirm('確認刪除該單位？');"
                                            TabIndex="65" Visible="false" />&nbsp;&nbsp;
                                        <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="bt" TabIndex="70" Visible="false"
                                            OnClientClick="return CheckValidate();" />&nbsp;&nbsp;
                                        <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="bt" TabIndex="75" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <asp:HiddenField ID="hidAddNodeValuePath" runat="server" />
    <asp:HiddenField ID="hidOperateFlag" runat="server" />
    <asp:HiddenField ID="hidDdlBraDepNo" runat="server" />
    <asp:HiddenField ID = "hidScript" runat = "server" />
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
