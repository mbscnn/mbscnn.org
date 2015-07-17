<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="Signin.ascx.vb" Inherits="MBSC.UICtl.Signin" %>
<div style="clear: both;"></div>
<div class="status-msg-wrap">
    <div id="DIV_TAB_TITLE" class="status-msg-body">
        <asp:PlaceHolder ID="PLH_LOGOUT" runat="server" Visible="false">
            <span style="float: right">
                <span style="font-weight: bold">
                    <asp:Literal ID="LTL_MB_NAME" runat="server" />
                </span>
                <asp:Button ID="btLogout" runat="server" Text="登出" CssClass="mbscbt" />
            </span>
        </asp:PlaceHolder>

        <asp:Literal ID="LTL_TAB_TITLE" runat="server" />
    </div>
    <div class="status-msg-border">
        <div class="status-msg-bg">
            <div class="status-msg-hidden">
                <asp:Literal ID="LTL_TAB_TITLE_HID" runat="server" />
            </div>
        </div>
    </div>
</div>

<div style="clear: both;"></div>